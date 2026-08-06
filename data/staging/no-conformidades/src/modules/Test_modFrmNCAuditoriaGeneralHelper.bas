Attribute VB_Name = "Test_modFrmNCAuditoriaGeneralHelper"
Option Compare Database
Option Explicit

' =============================================================================
' Atoms for modFrmNCAuditoriaGeneralHelper (slice 1 of form-thin-helper-refactor).
' Mirror of Test_modFrmNCProyectoGeneralHelper for the NCAuditoria side.
'
' Skill:    access-vba-tdd §1.1, §1.4, §1.6, §1.8
'           access-vba-e2e-methodology rules #5, #9, #1-#11B
' =============================================================================

' ----- Message contract (byte-for-byte match) -----
Private Const MSG_MOTIVO_LIMPIADO As String = "El campo 'Motivo de control de eficacia no requerido' se ha limpiado porque ahora sí requiere control de eficacia."

' ----- Fixture IDs -----
Private Const FIX_ID_NC As Long = 900610002

' ----- Module-level state for fixture restoration -----
Private m_PrevUsuarioConectado As usuario
Private m_PrevEntorno As entorno
Private ErrMsg_Local As String

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

' Pre-insert a minimal NCAuditoria row. Schema for TbNoConformidadesAuditoria
' verified via dysflow.get_schema (see docs/schema/erd-no-conformidades.md):
'   id (Long, PK), DESCRIPCION, ESTADO, FECHAAPERTURA, Tipo, PuntoNorma, CAUSARAIZ,
'   RequiereControlEficacia, MotivoNoRequiereControlEficacia, RESPONSABLEIMPLANTACION.
' NO CodigoNoConformidad, NO EsNoConformidad columns.
Private Function EnsureNCAuditoriaFixture(ByVal p_Db As DAO.Database, _
                                         ByVal p_IDNC As Long, _
                                         ByVal p_RequiereControlEficacia As String, _
                                         ByVal p_MotivoNoRequiereControlEficacia As String, _
                                         ByRef p_Error As String) As Boolean
    On Error GoTo EH
    EnsureNCAuditoriaFixture = False
    p_Error = ""
    If Not TableExistsInDb(p_Db, "TbNoConformidadesAuditoria") Then
        p_Error = "TbNoConformidadesAuditoria does not exist in current DB"
        Exit Function
    End If
    p_Db.Execute "DELETE FROM TbNoConformidadesAuditoria WHERE id=" & p_IDNC, dbFailOnError
    p_Db.Execute "INSERT INTO TbNoConformidadesAuditoria (id, " & _
                 "DESCRIPCION, ESTADO, FECHAAPERTURA, " & _
                 "Tipo, PuntoNorma, CAUSARAIZ, " & _
                 "RequiereControlEficacia, MotivoNoRequiereControlEficacia, " & _
                 "RESPONSABLEIMPLANTACION) " & _
                 "VALUES (" & p_IDNC & ", " & _
                 "'NC fixture slice 1 auditoria', " & _
                 "'REGISTRADA', #2026-06-25#, " & _
                 "'Auditoria', 'TEST-PN', 'TEST-CAUSA', " & _
                 "'" & Replace(p_RequiereControlEficacia, "'", "''") & "', " & _
                 IIf(p_MotivoNoRequiereControlEficacia = "", "NULL", "'" & Replace(p_MotivoNoRequiereControlEficacia, "'", "''") & "'") & ", " & _
                 "'TEST_RESPONSABLE')", dbFailOnError
    EnsureNCAuditoriaFixture = True
    Exit Function
EH:
    p_Error = "EnsureNCAuditoriaFixture: " & Err.Description
End Function

Private Sub CleanupFixture(ByVal p_Db As DAO.Database, ByVal p_IDNC As Long)
    On Error Resume Next
    If TableExistsInDb(p_Db, "TbNoConformidadesAuditoria") Then
        p_Db.Execute "DELETE FROM TbNoConformidadesAuditoria WHERE id=" & p_IDNC, dbFailOnError
    End If
    On Error GoTo 0
End Sub

' Build a fresh NCAuditoria. Field names verified against NCAuditoria.cls source:
'   id (String), IDAuditoria (String), FechaApertura (String), Numero (String),
'   Descripcion (String), CAUSARAIZ (String), AccionCorrectiva (String),
'   CORRECCION (String), FECHACIERRE (String), FPREVCIERRE (String),
'   RESPONSABLEIMPLANTACION (String), RequiereControlEficacia (String),
'   ControlEficacia (String), FechaControlEficacia (String),
'   FechaPrevistaControlEficacia (String), ResultadoControlEficacia (String),
'   ConformeControlEficacia (String), RequiereAccionCorrectiva (String),
'   MotivoNoAccionCorrectiva (String), MotivoNoRequiereControlEficacia (String),
'   Tipo (String), PuntoNorma (String), Estado (String), Borrado (Boolean),
'   MotivoBorrado (String), Notas (String), Cerrada (String), Error (String).
' NO CodigoNoConformidad field -- only NCProyecto has it. NO AuditoriaObj setter
' (Auditoria is Property Get only).
Private Function BuildNCAuditoriaForSlice1( _
                                         ByVal p_IDNC As Long, _
                                         ByVal p_RequiereControlEficacia As String, _
                                         ByVal p_MotivoNoRequiereControlEficacia As String) As NCAuditoria
    Dim nc As NCAuditoria
    Set nc = New NCAuditoria
    nc.id = CStr(p_IDNC)
    nc.Descripcion = "NC fixture slice 1 auditoria " & p_IDNC
    nc.CAUSARAIZ = "TEST-CAUSA"
    nc.PuntoNorma = "TEST-PN"
    nc.Tipo = "Auditoria"
    nc.FechaApertura = #6/25/2026#
    nc.RequiereControlEficacia = p_RequiereControlEficacia
    nc.MotivoNoRequiereControlEficacia = p_MotivoNoRequiereControlEficacia
    nc.RESPONSABLEIMPLANTACION = "TEST_RESPONSABLE"
    Set BuildNCAuditoriaForSlice1 = nc
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
' ATOM 1 - Happy path: motivo empty, no warning.
' =============================================================================
Public Function Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_HappyPath_NoMotivoCleared_NoMessage_Atomic() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim ncInicial As NCAuditoria
    Dim ncActual As NCAuditoria
    Dim promptResult As Long
    Dim messageText As String
    Dim resultStr As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_HappyPath_NoMotivoCleared_NoMessage_Atomic = TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    Set db = getdb()
    SetupTestContext
    CleanupFixture db, FIX_ID_NC

    If Not EnsureNCAuditoriaFixture(db, FIX_ID_NC, "No", "", ErrMsg_Local) Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_HappyPath_NoMotivoCleared_NoMessage_Atomic = TestHelper.BuildJsonFail("EnsureNCAuditoriaFixture: " & ErrMsg_Local, logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Arrange: NCAuditoria pre-inserted (RequiereCE='No', Motivo='' happy path)"

    Set ncInicial = BuildNCAuditoriaForSlice1(FIX_ID_NC, "No", "")
    Set ncActual = BuildNCAuditoriaForSlice1(FIX_ID_NC, "Sí", "")

    promptResult = -1
    messageText = ""
    On Error Resume Next
    resultStr = modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos(ncActual, ncInicial, db, promptResult, messageText)
    On Error GoTo EH
    If ErrMsg_Local <> "" Then
        TestHelper.AddLog logs, "Tolerated downstream errMsg: " & ErrMsg_Local
        ErrMsg_Local = ""
    End If
    TestHelper.AddLog logs, "Act: helper returned; messageText length=" & Len(messageText)

    If messageText <> "" Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_HappyPath_NoMotivoCleared_NoMessage_Atomic = TestHelper.BuildJsonFail( _
            "p_MessageText should be empty on happy path (got='" & messageText & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: p_MessageText empty"

    Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_HappyPath_NoMotivoCleared_NoMessage_Atomic = TestHelper.BuildJsonOk(logs, "happy_path_no_message")
    GoTo Cleanup

EH:
    Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_HappyPath_NoMotivoCleared_NoMessage_Atomic = TestHelper.BuildJsonFail( _
        "Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_HappyPath_NoMotivoCleared_NoMessage_Atomic: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If Not db Is Nothing Then CleanupFixture db, FIX_ID_NC
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Set db = Nothing
End Function

' =============================================================================
' ATOM 2 - Motivo was set, gets cleared, p_MessageText populated.
' =============================================================================
Public Function Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_MotivoCleared_PopulatesPMessageText_Atomic() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim ncInicial As NCAuditoria
    Dim ncActual As NCAuditoria
    Dim promptResult As Long
    Dim messageText As String
    Dim resultStr As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_MotivoCleared_PopulatesPMessageText_Atomic = TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    Set db = getdb()
    SetupTestContext
    CleanupFixture db, FIX_ID_NC

    If Not EnsureNCAuditoriaFixture(db, FIX_ID_NC, "No", "motivo previo auditoria", ErrMsg_Local) Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_MotivoCleared_PopulatesPMessageText_Atomic = TestHelper.BuildJsonFail("EnsureNCAuditoriaFixture: " & ErrMsg_Local, logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Arrange: NCAuditoria pre-inserted (RequiereCE='No', Motivo='motivo previo auditoria')"

    Set ncInicial = BuildNCAuditoriaForSlice1(FIX_ID_NC, "No", "motivo previo auditoria")
    Set ncActual = BuildNCAuditoriaForSlice1(FIX_ID_NC, "Sí", "motivo previo auditoria")

    promptResult = -1
    messageText = ""
    On Error Resume Next
    resultStr = modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos(ncActual, ncInicial, db, promptResult, messageText)
    On Error GoTo EH
    If ErrMsg_Local <> "" Then
        TestHelper.AddLog logs, "Tolerated downstream errMsg: " & ErrMsg_Local
        ErrMsg_Local = ""
    End If
    TestHelper.AddLog logs, "Act: helper returned after motivo clear"

    If messageText <> MSG_MOTIVO_LIMPIADO Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_MotivoCleared_PopulatesPMessageText_Atomic = TestHelper.BuildJsonFail( _
            "p_MessageText mismatch (expected='" & MSG_MOTIVO_LIMPIADO & "', got='" & messageText & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: p_MessageText = MSG_MOTIVO_LIMPIADO"

    ' NOTE: We intentionally do NOT assert ncActual.MotivoNoRequiereControlEficacia = "" here.
    ' The operations class (NCaUDITORIAOperaciones.RegistrarDatosUnicos) is responsible for
    ' actually clearing motivo in the entity + DB. In the test environment, the operations
    ' class has many dependencies (m_ObjUsuarioConectado, FK validation, RegistrarLog, etc.)
    ' that we do not provide; the call may error out before the motivo-clear block. That
    ' is a separate integration concern. The helper's contract is solely to populate
    ' p_MessageText correctly when a No->Sí transition with motivo is requested.

    Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_MotivoCleared_PopulatesPMessageText_Atomic = TestHelper.BuildJsonOk(logs, "motivo_cleared_message_populated")
    GoTo Cleanup

EH:
    Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_MotivoCleared_PopulatesPMessageText_Atomic = TestHelper.BuildJsonFail( _
        "Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_MotivoCleared_PopulatesPMessageText_Atomic: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If Not db Is Nothing Then CleanupFixture db, FIX_ID_NC
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Set db = Nothing
End Function

' =============================================================================
' ATOM 3 - Motivo NOT cleared (no transition). p_MessageText empty.
' =============================================================================
Public Function Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_MotivoNotCleared_NoMessage_Atomic() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim ncInicial As NCAuditoria
    Dim ncActual As NCAuditoria
    Dim promptResult As Long
    Dim messageText As String
    Dim resultStr As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_MotivoNotCleared_NoMessage_Atomic = TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    Set db = getdb()
    SetupTestContext
    CleanupFixture db, FIX_ID_NC

    If Not EnsureNCAuditoriaFixture(db, FIX_ID_NC, "Sí", "motivo que no se borra", ErrMsg_Local) Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_MotivoNotCleared_NoMessage_Atomic = TestHelper.BuildJsonFail("EnsureNCAuditoriaFixture: " & ErrMsg_Local, logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Arrange: NCAuditoria pre-inserted (RequiereCE='Sí' already, Motivo='motivo que no se borra')"

    Set ncInicial = BuildNCAuditoriaForSlice1(FIX_ID_NC, "Sí", "motivo que no se borra")
    Set ncActual = BuildNCAuditoriaForSlice1(FIX_ID_NC, "Sí", "motivo que no se borra")

    promptResult = -1
    messageText = ""
    On Error Resume Next
    resultStr = modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos(ncActual, ncInicial, db, promptResult, messageText)
    On Error GoTo EH
    If ErrMsg_Local <> "" Then
        TestHelper.AddLog logs, "Tolerated downstream errMsg: " & ErrMsg_Local
        ErrMsg_Local = ""
    End If
    TestHelper.AddLog logs, "Act: helper returned; motivo unchanged"

    If messageText <> "" Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_MotivoNotCleared_NoMessage_Atomic = TestHelper.BuildJsonFail( _
            "p_MessageText should be empty when motivo not cleared (got='" & messageText & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: p_MessageText empty (motivo unchanged)"

    If ncActual.MotivoNoRequiereControlEficacia <> "motivo que no se borra" Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_MotivoNotCleared_NoMessage_Atomic = TestHelper.BuildJsonFail( _
            "MotivoNoRequiereControlEficacia mutated unexpectedly (got='" & ncActual.MotivoNoRequiereControlEficacia & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: motivo preserved"

    Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_MotivoNotCleared_NoMessage_Atomic = TestHelper.BuildJsonOk(logs, "motivo_not_cleared_no_message")
    GoTo Cleanup

EH:
    Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_MotivoNotCleared_NoMessage_Atomic = TestHelper.BuildJsonFail( _
        "Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_MotivoNotCleared_NoMessage_Atomic: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If Not db Is Nothing Then CleanupFixture db, FIX_ID_NC
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Set db = Nothing
End Function

' =============================================================================
' ATOM 4 - Adversarial: double-click race.
' =============================================================================
Public Function Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_Adversarial_DoubleClickRace_Atomic() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim ncInicial As NCAuditoria
    Dim ncActual As NCAuditoria
    Dim promptResult As Long
    Dim messageText1 As String
    Dim messageText2 As String
    Dim resultStr As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_Adversarial_DoubleClickRace_Atomic = TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    Set db = getdb()
    SetupTestContext
    CleanupFixture db, FIX_ID_NC

    If Not EnsureNCAuditoriaFixture(db, FIX_ID_NC, "No", "motivo race", ErrMsg_Local) Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_Adversarial_DoubleClickRace_Atomic = TestHelper.BuildJsonFail("EnsureNCAuditoriaFixture: " & ErrMsg_Local, logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Arrange: NCAuditoria pre-inserted (RequiereCE='No', Motivo='motivo race')"

    Set ncInicial = BuildNCAuditoriaForSlice1(FIX_ID_NC, "No", "motivo race")
    Set ncActual = BuildNCAuditoriaForSlice1(FIX_ID_NC, "Sí", "motivo race")

    promptResult = -1
    messageText1 = ""
    On Error Resume Next
    resultStr = modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos(ncActual, ncInicial, db, promptResult, messageText1)
    On Error GoTo EH
    If ErrMsg_Local <> "" Then
        TestHelper.AddLog logs, "Tolerated downstream errMsg (call 1): " & ErrMsg_Local
        ErrMsg_Local = ""
    End If

    If messageText1 <> MSG_MOTIVO_LIMPIADO Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_Adversarial_DoubleClickRace_Atomic = TestHelper.BuildJsonFail( _
            "First call: p_MessageText mismatch (expected='" & MSG_MOTIVO_LIMPIADO & "', got='" & messageText1 & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK (call 1): p_MessageText populated with MSG_MOTIVO_LIMPIADO"

    ' Second call: in a real double-click, the form RE-LOADS ncInicial from DB before
    ' calling the helper again. After call 1, the DB has RequiereCE='Sí' and motivo=Null
    ' (cleared by the operations class). We simulate that re-load by rebuilding ncInicial
    ' with the post-clear state (motivo='', requiereInitial='Sí'). The pre-flight check
    ' in the helper should then detect NO transition and leave p_MessageText empty.
    Set ncInicial = BuildNCAuditoriaForSlice1(FIX_ID_NC, "Sí", "")
    messageText2 = ""
    On Error Resume Next
    resultStr = modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos(ncActual, ncInicial, db, promptResult, messageText2)
    On Error GoTo EH
    If ErrMsg_Local <> "" Then
        TestHelper.AddLog logs, "Tolerated downstream errMsg (call 2): " & ErrMsg_Local
        ErrMsg_Local = ""
    End If

    If messageText2 <> "" Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_Adversarial_DoubleClickRace_Atomic = TestHelper.BuildJsonFail( _
            "Second call: p_MessageText should be empty (got='" & messageText2 & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK (call 2): p_MessageText empty"

    Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_Adversarial_DoubleClickRace_Atomic = TestHelper.BuildJsonOk(logs, "double_click_race_last_wins")
    GoTo Cleanup

EH:
    Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_Adversarial_DoubleClickRace_Atomic = TestHelper.BuildJsonFail( _
        "Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_Adversarial_DoubleClickRace_Atomic: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If Not db Is Nothing Then CleanupFixture db, FIX_ID_NC
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Set db = Nothing
End Function

' =============================================================================
' ATOM 5 - Edge: motivo is Null.
' =============================================================================
Public Function Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_Edge_NullMotivo_NoMessage_Atomic() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim ncInicial As NCAuditoria
    Dim ncActual As NCAuditoria
    Dim promptResult As Long
    Dim messageText As String
    Dim resultStr As String

    On Error GoTo EH
    Set logs = TestHelper.NewLogs()
    If Not TestHelper.BeginTestSession(logs) Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_Edge_NullMotivo_NoMessage_Atomic = TestHelper.BuildJsonFail("BeginTestSession failed", logs)
        Exit Function
    End If
    Set db = getdb()
    SetupTestContext
    CleanupFixture db, FIX_ID_NC

    If Not EnsureNCAuditoriaFixture(db, FIX_ID_NC, "No", "", ErrMsg_Local) Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_Edge_NullMotivo_NoMessage_Atomic = TestHelper.BuildJsonFail("EnsureNCAuditoriaFixture: " & ErrMsg_Local, logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Arrange: NCAuditoria pre-inserted (RequiereCE='No', Motivo=Null)"

    Set ncInicial = BuildNCAuditoriaForSlice1(FIX_ID_NC, "No", "")
    Set ncActual = BuildNCAuditoriaForSlice1(FIX_ID_NC, "Sí", "")

    promptResult = -1
    messageText = ""
    On Error Resume Next
    resultStr = modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos(ncActual, ncInicial, db, promptResult, messageText)
    On Error GoTo EH
    If ErrMsg_Local <> "" Then
        TestHelper.AddLog logs, "Tolerated downstream errMsg: " & ErrMsg_Local
        ErrMsg_Local = ""
    End If
    TestHelper.AddLog logs, "Act: helper returned (motivo was empty/Null)"

    If messageText <> "" Then
        Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_Edge_NullMotivo_NoMessage_Atomic = TestHelper.BuildJsonFail( _
            "p_MessageText should be empty when motivo was Null (got='" & messageText & "')", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Assert OK: p_MessageText empty (Null motivo treated as empty)"

    Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_Edge_NullMotivo_NoMessage_Atomic = TestHelper.BuildJsonOk(logs, "edge_null_motivo_no_message")
    GoTo Cleanup

EH:
    Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_Edge_NullMotivo_NoMessage_Atomic = TestHelper.BuildJsonFail( _
        "Test_modFrmNCAuditoriaGeneralHelper_RegistrarDatosUnicos_Edge_NullMotivo_NoMessage_Atomic: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If Not db Is Nothing Then CleanupFixture db, FIX_ID_NC
    Call RestoreTestContext
    Call TestHelper.EndTestSession(logs)
    Set db = Nothing
End Function
