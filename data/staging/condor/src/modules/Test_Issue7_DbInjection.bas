Attribute VB_Name = "Test_Issue7_DbInjection"
Option Compare Database
Option Explicit

' ============================================================================
' Test_Issue7_DbInjection — focused test for GitHub Issue #7
'
' Closes: GH #7
' SDD:    staging-traceability-recovery
'
' Purpose
'   Verify that DatosPCSUBRepositorio.ActualizarAprobacionSuministrador honors
'   an injected DAO.Database parameter: it must (a) use the injected db for the
'   actual write, and (b) leave the caller's reference to the injected db
'   usable after the call returns.
'
' v2.4.2 migration: SetupSandbox (private) replaced with the canonical
'   TestHelper.BeginTestSession. Sandbox handle is opened explicitly against
'   m_BackendSandboxURL — TestHelper.GetTestDb() can return a stale cached
'   handle that points at the wrong db. CountPCSUBRows stays as a private
'   helper anchored to the injected db (no DCount, which always resolves
'   through CurrentDb() / the frontend and cannot see the backend tables).
' ============================================================================

Private Const TEST_ID_BASE As Long = 900000
Private Const ISSUE7_ID As Long = 900001

' ---------------------------------------------------------------------------
' TeardownAll: deletes all fixture rows in reverse FK order. Defensive
' (On Error Resume Next) so it never raises and never blocks the next
' test (TDD v2.4 §5 idempotency rule).
' ---------------------------------------------------------------------------
Private Sub TeardownAll(ByVal p_Db As DAO.Database)
    On Error Resume Next
    p_Db.Execute "DELETE FROM tbDatosPCSUB WHERE idSolicitud >= " & TEST_ID_BASE, dbFailOnError
    p_Db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud >= " & TEST_ID_BASE, dbFailOnError
    p_Db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente >= " & TEST_ID_BASE, dbFailOnError
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' SeedIssue7Fixture: minimal graph that satisfies the schema-first rule.
' Inserts ONLY the columns marked Required/NotNull in the real ERD:
'   - TbExpedientes: IDExpediente (PK)
'   - tbSolicitudes: idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud,
'     idEstadoInterno, fechaCreacion, usuarioCreacion, revisionCalidadEstado
'   - tbDatosPCSUB: idDatosPCSUB, idSolicitud, refContratoInspeccionOficial
' Returns the idSolicitud for the test to use.
' ---------------------------------------------------------------------------
Private Function SeedIssue7Fixture(ByVal p_Db As DAO.Database) As Long
    Dim sql As String

    sql = "INSERT INTO TbExpedientes (IDExpediente) VALUES (" & ISSUE7_ID & ")"
    p_Db.Execute sql, dbFailOnError

    sql = "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, " & _
          "codigoSolicitud, idEstadoInterno, fechaCreacion, usuarioCreacion, " & _
          "revisionCalidadEstado) VALUES (" & ISSUE7_ID & ", " & ISSUE7_ID & ", " & _
          "'PC_SUB', 'PCSUB-ISSUE7', 1, #2025-01-01 10:00:00#, 'TestIssue7', 'PENDIENTE')"
    p_Db.Execute sql, dbFailOnError

    sql = "INSERT INTO tbDatosPCSUB (idDatosPCSUB, idSolicitud, " & _
          "refContratoInspeccionOficial) VALUES (" & ISSUE7_ID & ", " & _
          ISSUE7_ID & ", 'REF-ISSUE7')"
    p_Db.Execute sql, dbFailOnError

    SeedIssue7Fixture = ISSUE7_ID
End Function

' ---------------------------------------------------------------------------
' CountPCSUBRows: helper for cardinality before/after. Returns the
' number of tbDatosPCSUB rows for a given idSolicitud. Uses db.OpenRecordset
' (NOT DCount) so the count is anchored to the injected backend db
' (TDD v2.4 §5.4 — db explicit, never CurrentDb fallback).
' ---------------------------------------------------------------------------
Private Function CountPCSUBRows(ByVal p_Db As DAO.Database, ByVal p_IdSol As Long) As Long
    Dim rs As DAO.Recordset
    On Error GoTo EH
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS n FROM tbDatosPCSUB WHERE idSolicitud = " & p_IdSol)
    CountPCSUBRows = Nz(rs!n, 0)
    rs.Close
    Set rs = Nothing
    Exit Function
EH:
    CountPCSUBRows = -1
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
End Function

' ============================================================================
' THE TEST
' ============================================================================
Public Function Test_PCSUB_ActualizarAprobacionSuministrador_PreservesInjectedDb() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim setupError As String
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim pcsub As DatosPCSUB
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    logs = TestHelper.NewLogsArray(8)

    ' --- Arrange 1: BeginTestSession (TDD v2.4 §3.6 production guard) ---
    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_PCSUB_ActualizarAprobacionSuministrador_PreservesInjectedDb = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    logs(0) = "1. BeginTestSession OK: m_TestingMode=True, sandbox validated"

    ' 1b. Open sandbox backend explicitly. TestHelper.GetTestDb() can return
    ' a stale cached handle that points at a different db; opening
    ' m_BackendSandboxURL directly guarantees the same file BeginTestSession
    ' just validated. This is the access-vba-tdd §5.4 rule.
    Set db = DBEngine.Workspaces(0).OpenDatabase( _
        m_BackendSandboxURL, _
        False, _
        False, _
        ";PWD=" & m_BackendSandboxPassword)

    ' --- Arrange 2: clean state and seed ---
    Call TeardownAll(db)
    idSol = SeedIssue7Fixture(db)
    logs(1) = "2. Seeded exact sandbox fixture: idSolicitud=" & idSol

    ' --- Arrange 3: cardinality BEFORE the mutation (TDD v2.4 §4.5) ---
    countBefore = CountPCSUBRows(db, idSol)
    If countBefore <> 1 Then
        logs(2) = "3. FAILED: expected exactly 1 seeded row, got " & countBefore
        logs(7) = "Test_PCSUB_ActualizarAprobacionSuministrador_PreservesInjectedDb: complete"
        Test_PCSUB_ActualizarAprobacionSuministrador_PreservesInjectedDb = _
            TestHelper.BuildJsonFail("fixture setup failed: expected 1 tbDatosPCSUB row, got " & countBefore, logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countBefore=" & countBefore & " (expected 1)"

    ' --- Act: call repository with injected db ---
    Set pcsub = New DatosPCSUB
    pcsub.idSolicitud = idSol
    pcsub.firmaOficinaTecnicaSubSuministradorNombre = "ISSUE7 OFICINA TECNICA"
    pcsub.firmaRepSubSuministradorNombre = "ISSUE7 REP SUB"

    DatosPCSUBRepositorio.ActualizarAprobacionSuministrador pcsub, db
    logs(3) = "4. Repository called with injected DAO.Database"

    ' --- Assert 1: cardinality AFTER the mutation (TDD v2.4 §4.5) ---
    countAfter = CountPCSUBRows(db, idSol)
    If countAfter <> 1 Then
        logs(4) = "5. FAILED: UPDATE changed row count: before=" & countBefore & " after=" & countAfter
        logs(7) = "Test_PCSUB_ActualizarAprobacionSuministrador_PreservesInjectedDb: complete"
        Test_PCSUB_ActualizarAprobacionSuministrador_PreservesInjectedDb = _
            TestHelper.BuildJsonFail("UPDATE changed row count: before=" & countBefore & " after=" & countAfter, logs)
        GoTo CleanExit
    End If
    logs(4) = "5. countAfter=" & countAfter & " (expected 1)"

    ' --- Assert 2: injected db reference remained usable (catches pre-fix bug) ---
    If db Is Nothing Then
        logs(5) = "6. FAILED: injected db reference was cleared (pre-fix bug)"
        logs(7) = "Test_PCSUB_ActualizarAprobacionSuministrador_PreservesInjectedDb: complete"
        Test_PCSUB_ActualizarAprobacionSuministrador_PreservesInjectedDb = _
            TestHelper.BuildJsonFail("injected db reference was cleared (pre-fix bug regression)", logs)
        GoTo CleanExit
    End If
    logs(5) = "6. injected db reference remained usable after call"

    ' --- Assert 3: signer fields persisted through the injected db ---
    Set rs = db.OpenRecordset( _
        "SELECT firmaOficinaTecnicaSubSuministradorNombre, firmaRepSubSuministradorNombre " & _
        "FROM tbDatosPCSUB WHERE idSolicitud = " & idSol, dbOpenSnapshot)
    If rs.EOF Then
        logs(6) = "7. FAILED: signer row not found after update"
        logs(7) = "Test_PCSUB_ActualizarAprobacionSuministrador_PreservesInjectedDb: complete"
        Test_PCSUB_ActualizarAprobacionSuministrador_PreservesInjectedDb = _
            TestHelper.BuildJsonFail("signer row not found after update", logs)
        GoTo CleanExit
    End If

    If Nz(rs!firmaOficinaTecnicaSubSuministradorNombre, "") <> "ISSUE7 OFICINA TECNICA" _
       Or Nz(rs!firmaRepSubSuministradorNombre, "") <> "ISSUE7 REP SUB" Then
        logs(6) = "7. FAILED: signer fields not persisted through injected db"
        logs(7) = "Test_PCSUB_ActualizarAprobacionSuministrador_PreservesInjectedDb: complete"
        Test_PCSUB_ActualizarAprobacionSuministrador_PreservesInjectedDb = _
            TestHelper.BuildJsonFail("signer fields not persisted through injected db", logs)
        GoTo CleanExit
    End If
    logs(6) = "7. signer fields persisted and re-read with same db"

    logs(7) = "8. Test_PCSUB_ActualizarAprobacionSuministrador_PreservesInjectedDb: PASS"
    Test_PCSUB_ActualizarAprobacionSuministrador_PreservesInjectedDb = _
        TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
    If Not db Is Nothing Then
        TeardownAll db
        db.Close
        Set db = Nothing
    End If
    ' Reset test session defensively (TDD v2.4 §3.7 exhaustive cleanup)
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(7) = "Test_PCSUB_ActualizarAprobacionSuministrador_PreservesInjectedDb: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_ActualizarAprobacionSuministrador_PreservesInjectedDb = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function
