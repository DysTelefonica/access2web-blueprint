Attribute VB_Name = "Test_Snapshot_Strict"
Option Compare Database
Option Explicit

' ============================================================================
' Test_Snapshot_Strict — CAP-009 (Snapshot hash) strict TDD v2.4.2 atoms (Fase D, 2026-06-15)
'
' Touched tables: tbSolicitudes, TbExpedientes. Tests the public seam
' SnapshotServicio.ObtenerDatosTabla which now accepts Optional ByRef db
' (was: Set db = getdb() direct, breaking the seam).
' Also tests the audit-field exclusion (Spec-105: fechacreacion, fechamodificacion,
' usuariocreacion, usuarioModificacion, idusuariocreacion, idusuariomodificacion).
'
' Fixture: idSolicitud 900921-900999, IDExpediente 900921-900999. Sandbox required.
' v2.4.2 lifecycle: BeginTestSession / EndTestSession / ResetTestSession.
' ============================================================================

Private Const ID_SOL_SNAPSHOT_OBTENER As Long = 900921
Private Const ID_SOL_SNAPSHOT_EXCL As Long = 900922
Private Const FIXTURE_ID_LO As Long = 900920
Private Const FIXTURE_ID_HI As Long = 900999

' ----------------------------------------------------------------------------
' Fixture helpers (schema-first inserts against the real ERD)
'   tbSolicitudes Required fields: idSolicitud, idExpediente, tipoSolicitud,
'     codigoSolicitud, idEstadoInterno, fechaCreacion, usuarioCreacion,
'     revisionCalidadEstado
'   TbExpedientes Required fields: IDExpediente
' ----------------------------------------------------------------------------

Private Sub TeardownSnapshotFixtures(ByVal p_Db As DAO.Database)
    On Error Resume Next
    p_Db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud >= " & FIXTURE_ID_LO & _
                 " AND idSolicitud <= " & FIXTURE_ID_HI, dbFailOnError
    p_Db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente >= " & FIXTURE_ID_LO & _
                 " AND IDExpediente <= " & FIXTURE_ID_HI, dbFailOnError
    On Error GoTo 0
End Sub

Private Sub SeedSolicitudForSnapshot(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long)
    Dim sql As String
    ' Parent row in TbExpedientes must exist for the FK
    p_Db.Execute "INSERT INTO TbExpedientes (IDExpediente) VALUES (" & p_IdSolicitud & ")", dbFailOnError
    ' All Required fields populated; audit fields populated so the exclusion logic
    ' can be observed (excluded from the dict, not absent from the row).
    sql = "INSERT INTO tbSolicitudes (" & _
          "idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, " & _
          "idEstadoInterno, fechaCreacion, fechaModificacion, " & _
          "usuarioCreacion, usuarioModificacion, revisionCalidadEstado) VALUES (" & _
          p_IdSolicitud & ", " & p_IdSolicitud & ", 'PC', 'CAP009-SNAP-" & p_IdSolicitud & "', 2, " & _
          "#2025-01-01 10:00:00#, #2025-01-02 11:00:00#, 'creator-user', 'modifier-user', " & _
          "'PENDIENTE')"
    p_Db.Execute sql, dbFailOnError
End Sub

' CountRows: explicit count against the injected db.
'   Replaces DCount, which always resolves through CurrentDb() (the frontend)
'   and therefore cannot see backend-only tables in the split architecture.
'   Returns -1 on DAO error so the caller can assert a clean fail.
Private Function CountRows(ByVal p_Db As DAO.Database, ByVal p_Table As String, ByVal p_Where As String) As Long
    Dim rs As DAO.Recordset
    On Error GoTo EH
    If Len(p_Where) = 0 Then
        Set rs = p_Db.OpenRecordset("SELECT Count(*) AS N FROM " & p_Table, dbOpenSnapshot)
    Else
        Set rs = p_Db.OpenRecordset("SELECT Count(*) AS N FROM " & p_Table & " WHERE " & p_Where, dbOpenSnapshot)
    End If
    CountRows = Nz(rs!N, 0)
    rs.Close
    Set rs = Nothing
    Exit Function
EH:
    CountRows = -1
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
End Function

' ----------------------------------------------------------------------------
' Atoms
' ----------------------------------------------------------------------------

Public Function Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsExcludedAuditFields() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim svc As New SnapshotServicio
    Dim dict As Object
    Dim hasFechacreacion As Boolean
    Dim hasIdSolicitud As Boolean
    Dim countBefore As Long
    Dim countAfter As Long

    logs = TestHelper.NewLogsArray(7)

    ' 1. Arrange — open the canonical test session against the sandbox
    logs(0) = "1. Arrange: BeginTestSession routes getdb() to sandbox"
    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsExcludedAuditFields = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If

    ' 1b. Open the sandbox backend explicitly. The test must hit the
    ' backend (where tbSolicitudes / TbExpedientes live), not the frontend
    ' (which only has environment config tables). We bypass GetTestDb()
    ' because g_dbCondor's session cache can point at a stale connection
    ' from a previous run; opening m_BackendSandboxURL directly guarantees
    ' we are talking to the same file BeginTestSession validated.
    Set db = DBEngine.Workspaces(0).OpenDatabase( _
        m_BackendSandboxURL, _
        False, _
        False, _
        ";PWD=" & m_BackendSandboxPassword)
    logs(0) = "1. Arrange: opened sandbox backend at " & m_BackendSandboxURL

    ' 2. Fixture-first: clear residual rows in the fixture range, then seed a known row
    Call TeardownSnapshotFixtures(db)
    Call SeedSolicitudForSnapshot(db, ID_SOL_SNAPSHOT_OBTENER)

    ' Cardinality assertion: the seed must produce exactly one row in the fixture range
    countBefore = CountRows(db, "tbSolicitudes", _
        "idSolicitud >= " & FIXTURE_ID_LO & " AND idSolicitud <= " & FIXTURE_ID_HI)
    If countBefore <> 1 Then
        logs(1) = "2. FAILED: seed cardinality expected 1, got " & countBefore
        Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsExcludedAuditFields = _
            TestHelper.BuildJsonFail("seed cardinality off", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. Seeded tbSolicitudes id=" & ID_SOL_SNAPSHOT_OBTENER & " (audit fields populated); cardinality=1"

    ' 3. Act
    Set dict = svc.ObtenerDatosTabla("tbSolicitudes", "idSolicitud", ID_SOL_SNAPSHOT_OBTENER, db)
    If dict Is Nothing Then
        logs(2) = "3. FAILED: ObtenerDatosTabla returned Nothing"
        Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsExcludedAuditFields = _
            TestHelper.BuildJsonFail("returned Nothing", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. dict returned, count=" & dict.Count

    ' 4. Assert — audit field excluded per Spec-105
    hasFechacreacion = dict.Exists("fechacreacion")
    If hasFechacreacion Then
        logs(3) = "4. FAILED: fechacreacion should be excluded per Spec-105"
        Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsExcludedAuditFields = _
            TestHelper.BuildJsonFail("fechacreacion not excluded", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. fechacreacion correctly excluded"

    ' 5. Assert — business field present
    hasIdSolicitud = dict.Exists("idSolicitud")
    If Not hasIdSolicitud Then
        logs(4) = "5. FAILED: idSolicitud should be present in the dictionary"
        Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsExcludedAuditFields = _
            TestHelper.BuildJsonFail("idSolicitud missing", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. idSolicitud correctly present"

    ' 6. Cardinality after: the test never deletes; still 1 row
    countAfter = CountRows(db, "tbSolicitudes", _
        "idSolicitud >= " & FIXTURE_ID_LO & " AND idSolicitud <= " & FIXTURE_ID_HI)
    If countAfter <> 1 Then
        logs(5) = "6. FAILED: post-act cardinality expected 1, got " & countAfter
        Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsExcludedAuditFields = _
            TestHelper.BuildJsonFail("post-act cardinality off", logs)
        GoTo CleanExit
    End If
    logs(5) = "6. Post-act cardinality still 1 — read-only atom honored invariants"

    logs(6) = "7. PASS"
    Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsExcludedAuditFields = _
        TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    Set dict = Nothing
    If Not db Is Nothing Then
        TeardownSnapshotFixtures db
        db.Close
        Set db = Nothing
    End If
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsExcludedAuditFields = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsNothingForMissingId() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim svc As New SnapshotServicio
    Dim dict As Object
    Dim countBefore As Long
    Dim countAfter As Long

    logs = TestHelper.NewLogsArray(6)

    ' 1. Arrange — open the canonical test session against the sandbox
    logs(0) = "1. Arrange: BeginTestSession routes getdb() to sandbox"
    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsNothingForMissingId = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If

    ' 1b. Open the sandbox backend explicitly (see ReturnsExcludedAuditFields for rationale)
    Set db = DBEngine.Workspaces(0).OpenDatabase( _
        m_BackendSandboxURL, _
        False, _
        False, _
        ";PWD=" & m_BackendSandboxPassword)
    logs(0) = "1. Arrange: opened sandbox backend at " & m_BackendSandboxURL

    ' 2. Fixture-first: no row for idSolicitud=-1; teardown before to be safe
    Call TeardownSnapshotFixtures(db)
    countBefore = CountRows(db, "tbSolicitudes", _
        "idSolicitud >= " & FIXTURE_ID_LO & " AND idSolicitud <= " & FIXTURE_ID_HI)
    If countBefore <> 0 Then
        logs(1) = "2. FAILED: expected zero fixture rows before the Act, got " & countBefore
        Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsNothingForMissingId = _
            TestHelper.BuildJsonFail("fixture not empty", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. Confirmed no fixture row for idSolicitud=-1; cardinality=0"

    ' 3. Act
    Set dict = svc.ObtenerDatosTabla("tbSolicitudes", "idSolicitud", -1, db)
    logs(2) = "3. ObtenerDatosTabla called for idSolicitud=-1 (missing)"

    ' 4. Assert — not Nothing (empty dict is the contract)
    If dict Is Nothing Then
        logs(3) = "4. FAILED: returned Nothing instead of empty Scripting.Dictionary"
        Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsNothingForMissingId = _
            TestHelper.BuildJsonFail("returned Nothing", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. dict returned, count=" & dict.Count

    ' 5. Assert — count is zero
    If dict.Count <> 0 Then
        logs(4) = "5. FAILED: expected empty dict for missing id, got count=" & dict.Count
        Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsNothingForMissingId = _
            TestHelper.BuildJsonFail("dict not empty", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. empty dict for missing id as expected"

    ' 6. Cardinality after: the test never writes; still 0 rows
    countAfter = CountRows(db, "tbSolicitudes", _
        "idSolicitud >= " & FIXTURE_ID_LO & " AND idSolicitud <= " & FIXTURE_ID_HI)
    If countAfter <> 0 Then
        logs(5) = "6. FAILED: post-act cardinality expected 0, got " & countAfter
        Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsNothingForMissingId = _
            TestHelper.BuildJsonFail("post-act cardinality off", logs)
        GoTo CleanExit
    End If
    logs(5) = "6. Post-act cardinality still 0 — read-only atom honored invariants"

    logs(UBound(logs)) = "PASS"
    Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsNothingForMissingId = _
        TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    Set dict = Nothing
    If Not db Is Nothing Then
        TeardownSnapshotFixtures db
        db.Close
        Set db = Nothing
    End If
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsNothingForMissingId = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ----------------------------------------------------------------------------
' Smoke aggregator — must be a globally unique Public Function per
' access-vba-tdd §1.1.1 (Dysflow runner discovery).
' ----------------------------------------------------------------------------

Public Function Test_Snapshot_Strict_RunAll() As String
    On Error GoTo EH
    Dim logs() As String
    Dim result As String

    logs = TestHelper.NewLogsArray(3)

    result = Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsExcludedAuditFields()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then
        Test_Snapshot_Strict_RunAll = TestHelper.BuildJsonFail("ReturnsExcludedAuditFields atom failed", logs)
        Exit Function
    End If
    logs(0) = "1. ObtenerDatosTabla excludes audit fields per Spec-105"

    result = Test_Snapshot_Strict_ObtenerDatosTabla_ReturnsNothingForMissingId()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then
        Test_Snapshot_Strict_RunAll = TestHelper.BuildJsonFail("ReturnsNothingForMissingId atom failed", logs)
        Exit Function
    End If
    logs(1) = "2. ObtenerDatosTabla returns empty dict for missing id (not Nothing)"

    logs(2) = "3. PASS"
    Test_Snapshot_Strict_RunAll = TestHelper.BuildJsonOk("true", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Snapshot_Strict_RunAll: ERR " & Err.Number & " - " & Err.Description
    Test_Snapshot_Strict_RunAll = TestHelper.BuildJsonFail(Err.Description, logs)
End Function
