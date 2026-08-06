Attribute VB_Name = "Test_Logs_Strict"
Option Compare Database
Option Explicit

' ============================================================================
' Test_Logs_Strict — CAP-009 Logs strict TDD v2.4.2 atoms (Fase B5, 2026-06-15)
'
' Touched tables: tbLogCambios, tbLogErrores. LogServicio writes only.
' Read + write via explicit DAO.Database injection. Cardinalidad on mutations.
' Fixture IDs in 900620-900699.
'
' v2.4.2 migration: SetupLogsSandbox (private) replaced with the canonical
'   TestHelper.BeginTestSession(logs, errMsg). TestHelper.GetTestDb()
'   replaced with explicit OpenSandboxForTest() helper. Production
'   environment globals (TempVars("DatosEnLocal"), m_URLRutaAplicacionLocal,
'   m_ObjEntorno, m_ObjUsuarioActivo) re-applied via ApplyProdEnvForTest
'   after BeginTestSession (EVE() does not set them for the production
'   guard path; the service still depends on them).
' v2.4.3 consolidation: ApplyProdEnvForTest / EnsureFolder / CleanupTempRoot
'   (private) replaced with TestHelper.SetupProdGlobalsForTest /
'   TestHelper.CleanupProdTempRoot canónicos.
' ============================================================================

Private Const TEST_ID_BASE As Long = 900620
Private Const TEST_ID_TOP As Long = 900699

Private Const FILE_PREFIX As String = "CAP009-STRICT-"

' ----------------------------------------------------------------------------
' Fixture helpers
' ----------------------------------------------------------------------------

Private Sub TeardownFixtures(ByVal p_Db As DAO.Database)
    On Error Resume Next
    p_Db.Execute "DELETE FROM tbLogCambios WHERE registro >= " & TEST_ID_BASE & " AND registro <= " & TEST_ID_TOP, dbFailOnError
    p_Db.Execute "DELETE FROM tbLogErrores WHERE idLogError >= " & TEST_ID_BASE & " AND idLogError <= " & TEST_ID_TOP, dbFailOnError
    On Error GoTo 0
End Sub

' CountLogCambiosByRegistro: explicit count against the injected db.
'   Returns -1 on DAO error so the caller can assert a clean fail.
Private Function CountLogCambiosByRegistro(ByVal p_Db As DAO.Database, ByVal p_Registro As Long) As Long
    Dim rs As DAO.Recordset
    On Error GoTo EH
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS n FROM tbLogCambios WHERE registro=" & p_Registro, dbOpenSnapshot)
    If Not rs.EOF Then CountLogCambiosByRegistro = Nz(rs!n, 0)
    rs.Close
    Set rs = Nothing
    Exit Function
EH:
    CountLogCambiosByRegistro = -1
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
End Function

Private Function GetLogCambioCampo(ByVal p_Db As DAO.Database, ByVal p_Registro As Long) As String
    Dim rs As DAO.Recordset
    On Error GoTo EH
    Set rs = p_Db.OpenRecordset("SELECT campo FROM tbLogCambios WHERE registro=" & p_Registro, dbOpenSnapshot)
    If Not rs.EOF Then GetLogCambioCampo = Nz(rs!campo, "")
    rs.Close
    Set rs = Nothing
    Exit Function
EH:
    GetLogCambioCampo = ""
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
End Function

Private Function GetLogCambioValorNuevo(ByVal p_Db As DAO.Database, ByVal p_Registro As Long) As String
    Dim rs As DAO.Recordset
    On Error GoTo EH
    Set rs = p_Db.OpenRecordset("SELECT valorNuevo FROM tbLogCambios WHERE registro=" & p_Registro, dbOpenSnapshot)
    If Not rs.EOF Then GetLogCambioValorNuevo = Nz(rs!valorNuevo, "")
    rs.Close
    Set rs = Nothing
    Exit Function
EH:
    GetLogCambioValorNuevo = ""
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
End Function

' CountLogErroresByIdLogError: explicit count for a specific log error row.
Private Function CountLogErroresByIdLogError(ByVal p_Db As DAO.Database, ByVal p_IdLogError As Long) As Long
    Dim rs As DAO.Recordset
    On Error GoTo EH
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS n FROM tbLogErrores WHERE idLogError=" & p_IdLogError, dbOpenSnapshot)
    If Not rs.EOF Then CountLogErroresByIdLogError = Nz(rs!n, 0)
    rs.Close
    Set rs = Nothing
    Exit Function
EH:
    CountLogErroresByIdLogError = -1
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
End Function

' CountLogErroresGlobal: total rows in tbLogErrores. Used to detect that
' the fire-and-forget RegistrarError() actually wrote a row.
Private Function CountLogErroresGlobal(ByVal p_Db As DAO.Database) As Long
    Dim rs As DAO.Recordset
    On Error GoTo EH
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS n FROM tbLogErrores", dbOpenSnapshot)
    If Not rs.EOF Then CountLogErroresGlobal = Nz(rs!n, 0)
    rs.Close
    Set rs = Nothing
    Exit Function
EH:
    CountLogErroresGlobal = -1
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
End Function

Private Function GetMaxLogErrorId(ByVal p_Db As DAO.Database) As Long
    Dim rs As DAO.Recordset
    On Error GoTo EH
    Set rs = p_Db.OpenRecordset("SELECT MAX(idLogError) AS maxId FROM tbLogErrores", dbOpenSnapshot)
    If Not rs.EOF Then GetMaxLogErrorId = Nz(rs!maxId, 0)
    rs.Close
    Set rs = Nothing
    Exit Function
EH:
    GetMaxLogErrorId = 0
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
End Function

Private Function GetMaxLogCambioId(ByVal p_Db As DAO.Database) As Long
    Dim rs As DAO.Recordset
    On Error GoTo EH
    Set rs = p_Db.OpenRecordset("SELECT MAX(idLogCambio) AS maxId FROM tbLogCambios", dbOpenSnapshot)
    If Not rs.EOF Then GetMaxLogCambioId = Nz(rs!maxId, 0)
    rs.Close
    Set rs = Nothing
    Exit Function
EH:
    GetMaxLogCambioId = 0
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
End Function

' ----------------------------------------------------------------------------
' OpenSandboxForTest: opens m_BackendSandboxURL (validated by BeginTestSession)
'   as a fresh DAO.Database. Returns the handle; caller must close in
'   CleanExit. Bypasses TestHelper.GetTestDb() to avoid stale g_dbCondor
'   session cache (TDD v2.4 §5.4).
' ----------------------------------------------------------------------------
Private Function OpenSandboxForTest() As DAO.Database
    Set OpenSandboxForTest = DBEngine.Workspaces(0).OpenDatabase( _
        m_BackendSandboxURL, _
        False, _
        False, _
        ";PWD=" & m_BackendSandboxPassword)
End Function

' ApplyProdEnvForTest + EnsureFolder + CleanupTempRoot migrados a
' TestHelper.SetupProdGlobalsForTest / TestHelper.CleanupProdTempRoot
' (v2.4.3). Cada test ahora llama directamente al helper canónico.

' ----------------------------------------------------------------------------
' Atoms
' ----------------------------------------------------------------------------

Public Function Test_Logs_Strict_RegistrarCambio_InsertsRowWithExpectedFields() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New LogServicio
    Dim registroId As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim persistedCampo As String
    Dim persistedValorNuevo As String

    logs = TestHelper.NewLogsArray(6)

    ' 1. Arrange — BeginTestSession + open sandbox + apply prod env
    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_Logs_Strict_RegistrarCambio_InsertsRowWithExpectedFields = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Set db = OpenSandboxForTest()
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap009_logs", "CAP009 Test User")
    logs(0) = "1. BeginTestSession OK; sandbox opened; prod env applied"

    ' 2. Fixture-first — clear residual log rows in the test range
    Call TeardownFixtures(db)

    ' 3. Cardinality BEFORE
    registroId = 900631
    countBefore = CountLogCambiosByRegistro(db, registroId)
    logs(1) = "2. countBefore=" & countBefore

    ' 4. Act
    svc.RegistrarCambio "tbSolicitudes", registroId, "idEstadoInterno", CLng(2), CLng(8), "UPDATE", db
    logs(2) = "3. RegistrarCambio called with explicit DAO.Database"

    ' 5. Cardinality AFTER
    countAfter = CountLogCambiosByRegistro(db, registroId)
    If countAfter <> countBefore + 1 Then
        logs(3) = "4. FAILED: cardinality wrong: before=" & countBefore & " after=" & countAfter
        Test_Logs_Strict_RegistrarCambio_InsertsRowWithExpectedFields = _
            TestHelper.BuildJsonFail("cardinality wrong", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=" & countAfter

    ' 6. Assert fields persisted
    persistedCampo = GetLogCambioCampo(db, registroId)
    persistedValorNuevo = GetLogCambioValorNuevo(db, registroId)
    logs(4) = "5. persisted: campo=[" & persistedCampo & "] valorNuevo=[" & persistedValorNuevo & "]"

    If persistedCampo <> "idEstadoInterno" Then
        logs(5) = "6. FAILED: campo mismatch: [" & persistedCampo & "]"
        Test_Logs_Strict_RegistrarCambio_InsertsRowWithExpectedFields = _
            TestHelper.BuildJsonFail("campo mismatch", logs)
        GoTo CleanExit
    End If
    If persistedValorNuevo <> "8" Then
        logs(5) = "6. FAILED: valorNuevo mismatch: [" & persistedValorNuevo & "]"
        Test_Logs_Strict_RegistrarCambio_InsertsRowWithExpectedFields = _
            TestHelper.BuildJsonFail("valorNuevo mismatch", logs)
        GoTo CleanExit
    End If
    logs(5) = "6. campo and valorNuevo persisted correctly"

    Test_Logs_Strict_RegistrarCambio_InsertsRowWithExpectedFields = _
        TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then
        TeardownFixtures db
        db.Close
        Set db = Nothing
    End If
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Logs_Strict_RegistrarCambio_InsertsRowWithExpectedFields = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Logs_Strict_RegistrarError_InsertsRowWithExpectedFields() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New LogServicio
    Dim countBefore As Long
    Dim countAfter As Long
    Dim lastIdLogError As Long
    Dim rs As DAO.Recordset

    logs = TestHelper.NewLogsArray(7)

    ' 1. Arrange
    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_Logs_Strict_RegistrarError_InsertsRowWithExpectedFields = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Set db = OpenSandboxForTest()
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap009_logs", "CAP009 Test User")
    logs(0) = "1. BeginTestSession OK; sandbox opened; prod env applied"

    Call TeardownFixtures(db)

    ' 2. Cardinality BEFORE
    countBefore = CountLogErroresGlobal(db)
    logs(1) = "2. countBefore=" & countBefore

    ' 3. Act
    ' Firma real: RegistrarError(numeroError, descripcion, modulo, procedimiento).
    ' NO tiene parametro db: el servicio llama LogErrorRepositorio.GuardarError sin inyectar.
    svc.RegistrarError 513, "descripcion test", "Test_Logs_Strict", "RegistrarError_InsertsRow"
    logs(2) = "3. RegistrarError called (no db param, fire-and-forget per service contract)"

    ' 4. Cardinality AFTER
    countAfter = CountLogErroresGlobal(db)
    If countAfter <> countBefore + 1 Then
        logs(3) = "4. FAILED: cardinality wrong: before=" & countBefore & " after=" & countAfter
        Test_Logs_Strict_RegistrarError_InsertsRowWithExpectedFields = _
            TestHelper.BuildJsonFail("cardinality wrong", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=" & countAfter

    ' 5. Localizar la nueva fila por el idLogError maximo (el servicio calcula MAX+1)
    lastIdLogError = GetMaxLogErrorId(db)
    logs(4) = "5. lastIdLogError=" & lastIdLogError

    ' 6. Assert fields persisted
    Set rs = db.OpenRecordset("SELECT modulo, procedimiento, numeroError, descripcionError FROM tbLogErrores WHERE idLogError=" & lastIdLogError, dbOpenSnapshot)
    If rs.EOF Then
        logs(5) = "6. FAILED: row not found for lastIdLogError"
        Test_Logs_Strict_RegistrarError_InsertsRowWithExpectedFields = _
            TestHelper.BuildJsonFail("row not found", logs)
        GoTo CleanExit
    End If
    logs(5) = "6. row found: modulo=[" & Nz(rs!modulo, "") & "] procedimiento=[" & Nz(rs!procedimiento, "") & "] numErr=" & Nz(rs!numeroError, 0)

    If Nz(rs!modulo, "") <> "Test_Logs_Strict" Then
        logs(6) = "7. FAILED: modulo mismatch"
        Test_Logs_Strict_RegistrarError_InsertsRowWithExpectedFields = _
            TestHelper.BuildJsonFail("modulo mismatch", logs)
        GoTo CleanExit
    End If
    If Nz(rs!procedimiento, "") <> "RegistrarError_InsertsRow" Then
        logs(6) = "7. FAILED: procedimiento mismatch"
        Test_Logs_Strict_RegistrarError_InsertsRowWithExpectedFields = _
            TestHelper.BuildJsonFail("procedimiento mismatch", logs)
        GoTo CleanExit
    End If
    If Nz(rs!numeroError, 0) <> 513 Then
        logs(6) = "7. FAILED: numeroError mismatch: " & Nz(rs!numeroError, 0)
        Test_Logs_Strict_RegistrarError_InsertsRowWithExpectedFields = _
            TestHelper.BuildJsonFail("numeroError mismatch", logs)
        GoTo CleanExit
    End If
    logs(6) = "7. modulo, procedimiento, numeroError persisted correctly"

    Test_Logs_Strict_RegistrarError_InsertsRowWithExpectedFields = _
        TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
    If Not db Is Nothing Then
        TeardownFixtures db
        db.Close
        Set db = Nothing
    End If
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Logs_Strict_RegistrarError_InsertsRowWithExpectedFields = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Logs_Strict_LogCambioRepositorio_EliminarPorSolicitud_DeletesAllRows() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New LogServicio
    Dim registroId As Long
    Dim countBefore As Long
    Dim countAfter As Long

    logs = TestHelper.NewLogsArray(5)

    ' 1. Arrange
    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_Logs_Strict_LogCambioRepositorio_EliminarPorSolicitud_DeletesAllRows = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Set db = OpenSandboxForTest()
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap009_logs", "CAP009 Test User")
    logs(0) = "1. BeginTestSession OK; sandbox opened; prod env applied"

    Call TeardownFixtures(db)

    ' 2. Seed — 2 log cambios for the same registro
    registroId = 900651
    svc.RegistrarCambio "tbSolicitudes", registroId, "idEstadoInterno", CLng(2), CLng(3), "UPDATE", db
    svc.RegistrarCambio "tbSolicitudes", registroId, "idEstadoInterno", CLng(3), CLng(4), "UPDATE", db

    countBefore = CountLogCambiosByRegistro(db, registroId)
    If countBefore <> 2 Then
        logs(1) = "2. FAILED: seed cardinalidad expected 2, got " & countBefore
        Test_Logs_Strict_LogCambioRepositorio_EliminarPorSolicitud_DeletesAllRows = _
            TestHelper.BuildJsonFail("seed cardinalidad", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. countBefore=2"

    ' 3. Act
    LogCambioRepositorio.EliminarPorSolicitud registroId, db
    logs(2) = "3. EliminarPorSolicitud called with explicit DAO.Database"

    ' 4. Cardinality AFTER — expect 0
    countAfter = CountLogCambiosByRegistro(db, registroId)
    If countAfter <> 0 Then
        logs(3) = "4. FAILED: expected count=0 after delete, got " & countAfter
        Test_Logs_Strict_LogCambioRepositorio_EliminarPorSolicitud_DeletesAllRows = _
            TestHelper.BuildJsonFail("delete cardinalidad", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=0 as expected (all log cambios for registro=" & registroId & " deleted)"

    Test_Logs_Strict_LogCambioRepositorio_EliminarPorSolicitud_DeletesAllRows = _
        TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then
        TeardownFixtures db
        db.Close
        Set db = Nothing
    End If
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Logs_Strict_LogCambioRepositorio_EliminarPorSolicitud_DeletesAllRows = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ----------------------------------------------------------------------------
' Smoke aggregator — globally unique Public Function per §1.1.1
' ----------------------------------------------------------------------------
Public Function Test_Logs_Strict_RunAll() As String
    On Error GoTo EH
    Dim logs() As String
    Dim result As String

    logs = TestHelper.NewLogsArray(4)

    result = Test_Logs_Strict_RegistrarCambio_InsertsRowWithExpectedFields()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(0) = "1. RegistrarCambio inserts row with expected fields"

    result = Test_Logs_Strict_RegistrarError_InsertsRowWithExpectedFields()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(1) = "2. RegistrarError inserts row with expected fields"

    result = Test_Logs_Strict_LogCambioRepositorio_EliminarPorSolicitud_DeletesAllRows()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(2) = "3. EliminarPorSolicitud deletes all log cambios for the given registro"

    logs(3) = "4. PASS"
    Test_Logs_Strict_RunAll = TestHelper.BuildJsonOk("true", logs)
    Exit Function

Failed:
    Test_Logs_Strict_RunAll = TestHelper.BuildJsonFail("one CAP-009 logs atom failed", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Logs_Strict_RunAll: ERR " & Err.Number & " - " & Err.Description
    Test_Logs_Strict_RunAll = TestHelper.BuildJsonFail(Err.Description, logs)
End Function
