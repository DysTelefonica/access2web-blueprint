Attribute VB_Name = "Test_Issue19_EstadoValidacion"
Option Compare Database
Option Explicit

' ============================================================================
' Test_Issue19_EstadoValidacion — focused tests for #19 compile blocker and
'                                  #18 transactional isolation
'
' Closes: GH #19 (estadoValidacion undefined) and verifies the transactional
'          isolation fix in DatosCDCASUBServicio for GH #18.
' SDD:    fix-cdcASUB-servicio-transactional-isolation
'
' v2.4.2 migration: SetupSandbox (private, inlined) replaced with the
'   canonical TestHelper.BeginTestSession(logs, errMsg). Sandbox handle is
'   opened explicitly against m_BackendSandboxURL — TestHelper.GetTestDb()
'   can return a stale cached handle pointing at the wrong db. Inline
'   m_TestingMode / m_BackendSandboxURL cleanup is replaced by the
'   exhaustive TestHelper.ResetTestSession. CountCDCASUBRows stays as a
'   private helper anchored to the injected db (NOT DCount, which always
'   resolves through CurrentDb() / the frontend).
'
'   The two wrapper atoms (Test_Issue19_InValidation,
'   Test_Issue19_NotInValidation) are kept because they are already unique
'   global Public Functions per skill §1.1.1 and the existing
'   tests.issue19.json manifest references them.
' ============================================================================

Private Const TEST_ID_BASE As Long = 900000
Private Const ISSUE19_ID_VALIDACION As Long = 900010
Private Const ISSUE19_ID_PREREGISTRO As Long = 900011

' ---------------------------------------------------------------------------
' TeardownAll: deletes all fixture rows in reverse FK order. Defensive
' (On Error Resume Next) so it never raises and never blocks the next
' test (TDD v2.4 §5 idempotency rule).
' ---------------------------------------------------------------------------
Private Sub TeardownAll(ByVal p_Db As DAO.Database)
    On Error Resume Next
    p_Db.Execute "DELETE FROM tbDatosCDCASUB WHERE idSolicitud >= " & TEST_ID_BASE, dbFailOnError
    p_Db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud >= " & TEST_ID_BASE, dbFailOnError
    p_Db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente >= " & TEST_ID_BASE, dbFailOnError
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' SeedIssue19Fixture: minimal graph that satisfies the schema-first rule.
' Inserts ONLY the columns marked Required/NotNull in the real ERD:
'   - TbExpedientes: IDExpediente (PK)
'   - tbSolicitudes: idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud,
'     idEstadoInterno, fechaCreacion, usuarioCreacion, revisionCalidadEstado
'   - tbDatosCDCASUB: idDatosCDCASUB, idSolicitud
' Deterministic fechaCreacion literal (not Now()) so the fixture is
' reproducible across runs.
'
' p_IdEstadoInterno: the idEstadoInterno to use. 5 = validation, 1 = preregistro.
' Returns the idSolicitud.
' ---------------------------------------------------------------------------
Private Function SeedIssue19Fixture(ByVal p_Db As DAO.Database, _
                                     ByVal p_IdEstadoInterno As Long, _
                                     ByVal p_FixtureId As Long) As Long
    Dim sql As String
    Dim idExp As Long
    Dim idSol As Long
    Dim idDatos As Long

    idExp = p_FixtureId
    idSol = p_FixtureId
    idDatos = p_FixtureId

    sql = "INSERT INTO TbExpedientes (IDExpediente) VALUES (" & idExp & ")"
    p_Db.Execute sql, dbFailOnError

    sql = "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, " & _
          "codigoSolicitud, idEstadoInterno, fechaCreacion, usuarioCreacion, " & _
          "revisionCalidadEstado) VALUES (" & idSol & ", " & idExp & ", " & _
          "'CD_CA_SUB', 'CDCASUB-ISSUE19', " & p_IdEstadoInterno & ", #2025-01-01 10:00:00#, " & _
          "'TestIssue19', 'PENDIENTE')"
    p_Db.Execute sql, dbFailOnError

    sql = "INSERT INTO tbDatosCDCASUB (idDatosCDCASUB, idSolicitud) VALUES (" & _
          idDatos & ", " & idSol & ")"
    p_Db.Execute sql, dbFailOnError

    SeedIssue19Fixture = idSol
End Function

' ---------------------------------------------------------------------------
' CountCDCASUBRows: helper for cardinality before/after. Returns the
' number of tbDatosCDCASUB rows for a given idSolicitud. Uses db.OpenRecordset
' (NOT DCount) so the count is anchored to the injected backend db
' (TDD v2.4 §5.4 — db explicit, never CurrentDb fallback).
' ---------------------------------------------------------------------------
Private Function CountCDCASUBRows(ByVal p_Db As DAO.Database, ByVal p_IdSol As Long) As Long
    Dim rs As DAO.Recordset
    On Error GoTo EH
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS n FROM tbDatosCDCASUB WHERE idSolicitud = " & p_IdSol)
    CountCDCASUBRows = Nz(rs!n, 0)
    rs.Close
    Set rs = Nothing
    Exit Function
EH:
    CountCDCASUBRows = -1
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
End Function

' ============================================================================
' Wrapper atoms. Both are unique global Public Functions (skill §1.1.1) and
' are referenced by tests/tests.issue19.json. They delegate to the real
' implementations below.
' ============================================================================

Public Function Test_Issue19_InValidation() As String
    Test_Issue19_InValidation = Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_InValidation_UsesTransaction()
End Function

Public Function Test_Issue19_NotInValidation() As String
    Test_Issue19_NotInValidation = Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_NotInValidation_DoesNotFail()
End Function

' ============================================================================
' TEST 1 — In validation case (idEstadoInterno = estadoValidacion = 5)
' ============================================================================

Public Function Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_InValidation_UsesTransaction() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim setupError As String
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim obj As DatosCDCASUB
    Dim servicio As DatosCDCASUBServicio
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    logs = TestHelper.NewLogsArray(8)

    ' --- Arrange 1: BeginTestSession (TDD v2.4 §3.6 production guard) ---
    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_InValidation_UsesTransaction = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    logs(0) = "1. BeginTestSession OK"

    ' 1b. Open sandbox backend explicitly. TestHelper.GetTestDb() can return
    ' a stale cached handle; opening m_BackendSandboxURL directly guarantees
    ' the same file BeginTestSession validated (TDD v2.4 §5.4).
    Set db = DBEngine.Workspaces(0).OpenDatabase( _
        m_BackendSandboxURL, _
        False, _
        False, _
        ";PWD=" & m_BackendSandboxPassword)

    ' --- Arrange 2: clean state and seed solicitud in VALIDATION state ---
    Call TeardownAll(db)
    idSol = SeedIssue19Fixture(db, estadoValidacion, ISSUE19_ID_VALIDACION)
    logs(1) = "2. Seeded solicitud with idEstadoInterno=estadoValidacion=" & estadoValidacion

    ' --- Arrange 3: cardinality BEFORE the mutation (TDD v2.4 §4.5) ---
    countBefore = CountCDCASUBRows(db, idSol)
    If countBefore <> 1 Then
        logs(2) = "3. FAILED: expected 1 seeded row, got " & countBefore
        Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_InValidation_UsesTransaction = _
            TestHelper.BuildJsonFail("fixture setup failed: expected 1 tbDatosCDCASUB row, got " & countBefore, logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countBefore=" & countBefore & " (expected 1)"

    ' --- Act: call service with injected db = Nothing ---
    Set obj = New DatosCDCASUB
    obj.idSolicitud = idSol
    obj.racCodigo = "ISSUE19-RAC-CODE"
    obj.racDecision = "APROBADO"
    obj.racNombre = "ISSUE19-RAC-NAME"

    ' DatosCDCASUBServicio is a class (.cls) without VB_PredeclaredId,
    ' so it MUST be instantiated before use. Calling it as a static
    ' module (DatosCDCASUBServicio.GuardarDictamenRAC) is a compile error.
    Set servicio = New DatosCDCASUBServicio
    servicio.GuardarDictamenRAC obj
    logs(3) = "4. GuardarDictamenRAC called with no db (solicitud in validation)"

    ' --- Assert 1: cardinality AFTER the mutation ---
    countAfter = CountCDCASUBRows(db, idSol)
    If countAfter <> 1 Then
        logs(4) = "5. FAILED: countBefore=" & countBefore & " countAfter=" & countAfter
        Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_InValidation_UsesTransaction = _
            TestHelper.BuildJsonFail("countAfter=" & countAfter & " (expected 1)", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. countAfter=" & countAfter & " (expected 1)"

    ' --- Assert 2: no Err object state ---
    If Err.Number <> 0 Then
        logs(5) = "6. FAILED: Err.Number=" & Err.Number & " Desc=" & Err.Description
        Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_InValidation_UsesTransaction = _
            TestHelper.BuildJsonFail("Err state after call: " & Err.Description, logs)
        GoTo CleanExit
    End If
    logs(5) = "6. No Err state after call"

    ' --- Assert 3: RAC fields persisted through the service-opened connection ---
    Set rs = db.OpenRecordset( _
        "SELECT racCodigo, racDecision, racNombre FROM tbDatosCDCASUB WHERE idSolicitud = " & idSol, _
        dbOpenSnapshot)
    If rs.EOF Then
        logs(6) = "7. FAILED: row missing after update"
        Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_InValidation_UsesTransaction = _
            TestHelper.BuildJsonFail("row missing", logs)
        GoTo CleanExit
    End If
    If Nz(rs!racCodigo, "") <> "ISSUE19-RAC-CODE" Or Nz(rs!racDecision, "") <> "APROBADO" Or Nz(rs!racNombre, "") <> "ISSUE19-RAC-NAME" Then
        logs(6) = "7. FAILED: fields mismatch"
        Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_InValidation_UsesTransaction = _
            TestHelper.BuildJsonFail("RAC fields not persisted", logs)
        GoTo CleanExit
    End If
    rs.Close
    Set rs = Nothing
    logs(6) = "7. RAC fields persisted (racCodigo, racDecision, racNombre)"

    logs(7) = "8. Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_InValidation_UsesTransaction: PASS"
    Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_InValidation_UsesTransaction = _
        TestHelper.BuildJsonOk("invalidation_committed", logs)

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
    logs(7) = "Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_InValidation_UsesTransaction: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_InValidation_UsesTransaction = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ============================================================================
' TEST 2 — NOT in validation case (idEstadoInterno = estadoPreregistro = 1)
' ============================================================================
Public Function Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_NotInValidation_DoesNotFail() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim setupError As String
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim obj As DatosCDCASUB
    Dim servicio As DatosCDCASUBServicio
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    logs = TestHelper.NewLogsArray(8)

    ' --- Arrange 1: BeginTestSession ---
    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_NotInValidation_DoesNotFail = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    logs(0) = "1. BeginTestSession OK"

    ' 1b. Open sandbox backend explicitly (see InValidation test for rationale)
    Set db = DBEngine.Workspaces(0).OpenDatabase( _
        m_BackendSandboxURL, _
        False, _
        False, _
        ";PWD=" & m_BackendSandboxPassword)

    ' --- Arrange 2: seed solicitud in PREREGISTRO (NOT in validation) ---
    Call TeardownAll(db)
    idSol = SeedIssue19Fixture(db, 1, ISSUE19_ID_PREREGISTRO)
    logs(1) = "2. Seeded solicitud with idEstadoInterno=1 (NOT in validation)"

    ' --- Arrange 3: cardinality BEFORE ---
    countBefore = CountCDCASUBRows(db, idSol)
    If countBefore <> 1 Then
        logs(2) = "3. FAILED: expected 1 seeded row, got " & countBefore
        Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_NotInValidation_DoesNotFail = _
            TestHelper.BuildJsonFail("fixture setup failed: expected 1 tbDatosCDCASUB row, got " & countBefore, logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countBefore=" & countBefore & " (expected 1)"

    ' --- Act: call service with no db ---
    Set obj = New DatosCDCASUB
    obj.idSolicitud = idSol
    obj.racCodigo = "ISSUE19-NOTVAL-RAC-CODE"
    obj.racDecision = "APROBADO"
    obj.racNombre = "ISSUE19-NOTVAL-RAC-NAME"

    Set servicio = New DatosCDCASUBServicio
    servicio.GuardarDictamenRAC obj
    logs(3) = "4. GuardarDictamenRAC called with no db (solicitud NOT in validation)"

    ' --- Assert 1: cardinality AFTER ---
    countAfter = CountCDCASUBRows(db, idSol)
    If countAfter <> 1 Then
        logs(4) = "5. FAILED: countBefore=" & countBefore & " countAfter=" & countAfter
        Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_NotInValidation_DoesNotFail = _
            TestHelper.BuildJsonFail("countAfter=" & countAfter & " (expected 1)", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. countAfter=" & countAfter & " (expected 1)"

    ' --- Assert 2: no Err state ---
    If Err.Number <> 0 Then
        logs(5) = "6. FAILED: Err.Number=" & Err.Number & " Desc=" & Err.Description
        Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_NotInValidation_DoesNotFail = _
            TestHelper.BuildJsonFail("Err state after call: " & Err.Description, logs)
        GoTo CleanExit
    End If
    logs(5) = "6. No Err state after call"

    ' --- Assert 3: RAC fields persisted ---
    Set rs = db.OpenRecordset( _
        "SELECT racCodigo, racDecision, racNombre FROM tbDatosCDCASUB WHERE idSolicitud = " & idSol, _
        dbOpenSnapshot)
    If rs.EOF Then
        logs(6) = "7. FAILED: row missing after update"
        Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_NotInValidation_DoesNotFail = _
            TestHelper.BuildJsonFail("row missing", logs)
        GoTo CleanExit
    End If
    If Nz(rs!racCodigo, "") <> "ISSUE19-NOTVAL-RAC-CODE" Or Nz(rs!racDecision, "") <> "APROBADO" Or Nz(rs!racNombre, "") <> "ISSUE19-NOTVAL-RAC-NAME" Then
        logs(6) = "7. FAILED: fields mismatch"
        Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_NotInValidation_DoesNotFail = _
            TestHelper.BuildJsonFail("RAC fields not persisted", logs)
        GoTo CleanExit
    End If
    rs.Close
    Set rs = Nothing
    logs(6) = "7. RAC fields persisted"

    logs(7) = "8. Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_NotInValidation_DoesNotFail: PASS"
    Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_NotInValidation_DoesNotFail = _
        TestHelper.BuildJsonOk("not_in_validation_persisted", logs)

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
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(7) = "Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_NotInValidation_DoesNotFail: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_CDCASUBServicio_GuardarDictamenRAC_NotInValidation_DoesNotFail = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function
