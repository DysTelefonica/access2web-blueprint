Attribute VB_Name = "Test_CDCASUB_Strict"
Option Compare Database
Option Explicit

' ============================================================================
' Test_CDCASUB_Strict — CAP-005 CDCASUB strict TDD v2.4.2 atoms (Fase B2, 2026-06-15)
'
' Touched tables: tbSolicitudes, TbExpedientes, tbDatosCDCASUB.
' Read + write via explicit DAO.Database injection. Cardinalidad on
' mutations. Fixture IDs in 900810-900899.
'
' Note: EsParteTecnicaCompleta is exclusive to CDCASUB (does not exist on PC).
' EsAprobacionSuministradorCompleta requires firmaAprobacionRespIngenieriaNombre
' AND firmaAprobacionRespCalidadNombre (PC uses oficinaTecnica + repSuministrador).
' These divergences are the value of B2 vs B1.
'
' v2.4.2 migration: SetupCdcasubSandbox (private) replaced with the canonical
'   TestHelper.BeginTestSession(logs, errMsg). The hardened production guard
'   (UNC rejection, fingerprint check, FSO + DAO open) lives in
'   TestHelper.ForceLocalBackend and is now shared. TestHelper.GetTestDb()
'   replaced with explicit DBEngine.Workspaces(0).OpenDatabase(
'   m_BackendSandboxURL, ..., ";PWD=" & m_BackendSandboxPassword) so the
'   test is anchored to the same file BeginTestSession validated.
'
'   EVE() does not set m_ObjEntorno / m_ObjUsuarioActivo /
'   m_URLRutaAplicacionLocal / TempVars("DatosEnLocal") for the production
'   guard path, and the service depends on them. After BeginTestSession we
'   re-apply that exact setup so the service has what it needs.
' ============================================================================

Private Const TEST_ID_BASE As Long = 900810
Private Const TEST_ID_TOP As Long = 900899

Private Const ID_SOL_CDCASUB_PARTE As Long = 900811
Private Const ID_SOL_CDCASUB_SAD_CLASIF As Long = 900812
Private Const ID_SOL_CDCASUB_SAD_FIRMA As Long = 900813
Private Const ID_SOL_CDCASUB_RAC As Long = 900814
Private Const ID_SOL_CDCASUB_DECISION As Long = 900815
Private Const ID_SOL_CDCASUB_GENERALES As Long = 900816

Private Const FILE_PREFIX As String = "CAP005-STRICT-"

' ----------------------------------------------------------------------------
' Fixture helpers (schema-first inserts against the real ERD)
' ----------------------------------------------------------------------------

Private Sub TeardownFixtures(ByVal p_Db As DAO.Database)
    On Error Resume Next
    p_Db.Execute "DELETE FROM tbDatosCDCASUB WHERE idSolicitud >= " & TEST_ID_BASE & " AND idSolicitud <= " & TEST_ID_TOP, dbFailOnError
    p_Db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud >= " & TEST_ID_BASE & " AND idSolicitud <= " & TEST_ID_TOP, dbFailOnError
    p_Db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente >= " & TEST_ID_BASE & " AND IDExpediente <= " & TEST_ID_TOP, dbFailOnError
    On Error GoTo 0
End Sub

Private Sub SeedSolicitudCdcasub(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long, ByVal p_Codigo As String)
    On Error Resume Next
    p_Db.Execute "DELETE FROM tbDatosCDCASUB WHERE idSolicitud=" & p_IdSolicitud, dbFailOnError
    p_Db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud=" & p_IdSolicitud, dbFailOnError
    p_Db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & p_IdSolicitud, dbFailOnError
    On Error GoTo 0
    p_Db.Execute "INSERT INTO TbExpedientes (IDExpediente) VALUES (" & p_IdSolicitud & ")", dbFailOnError
    Dim sql As String
    sql = "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, " & _
          "idEstadoInterno, fechaCreacion, usuarioCreacion, revisionCalidadEstado) VALUES (" & _
          p_IdSolicitud & ", " & p_IdSolicitud & ", 'CDCA-SUB', '" & p_Codigo & "', 2, #2025-01-01 10:00:00#, " & _
          "'TestCdcasubStrict', 'PENDIENTE')"
    p_Db.Execute sql, dbFailOnError
End Sub

Private Sub SeedDatosCDCASUBParteTecnicaCompleta(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long)
    ' Seed a tbDatosCDCASUB row with the EXACT column names from ERD
    ' condor_datos.md §tbDatosCDCASUB. Used by EsParteTecnicaCompleta, which
    ' requires identificacionMaterial, causaNC, descripcionImpactoNC, and
    ' clasificacionNC in {MAYOR, MENOR} with esSubSuministradorAD non-null.
    Dim sql As String
    sql = "INSERT INTO tbDatosCDCASUB (idDatosCDCASUB, idSolicitud, " & _
          "refSubSuministrador, suministradorPrincipalNombreDir, " & _
          "subSuministradorNombreDir, requiereModificacionContrato, " & _
          "identificacionMaterial, causaNC, descripcionImpactoNC, " & _
          "clasificacionNC, esSubSuministradorAD) VALUES (" & _
          p_IdSolicitud & ", " & p_IdSolicitud & ", " & _
          "'REF-SUB-001', 'principal dir test', " & _
          "'sub dir test', False, " & _
          "'material test', 'causa test', 'impacto test', " & _
          "'MAYOR', True)"
    p_Db.Execute sql, dbFailOnError
End Sub

Private Sub SeedDatosCDCASUBParteTecnicaClasificacionInvalida(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long)
    ' Mismas columnas semilla que SeedDatosCDCASUBParteTecnicaCompleta
    ' pero con clasificacionNC='INVALID' para forzar EsParteTecnicaCompleta=False.
    Dim sql As String
    sql = "INSERT INTO tbDatosCDCASUB (idDatosCDCASUB, idSolicitud, " & _
          "refSubSuministrador, suministradorPrincipalNombreDir, " & _
          "subSuministradorNombreDir, requiereModificacionContrato, " & _
          "identificacionMaterial, causaNC, descripcionImpactoNC, " & _
          "clasificacionNC, esSubSuministradorAD) VALUES (" & _
          p_IdSolicitud & ", " & p_IdSolicitud & ", " & _
          "'REF-SUB-002', 'principal dir test', " & _
          "'sub dir test', False, " & _
          "'material test', 'causa test', 'impacto test', " & _
          "'INVALID', True)"
    p_Db.Execute sql, dbFailOnError
End Sub

Private Sub SeedDatosCDCASUBAprobacionSuministradorSinIngenieria(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long)
    ' Fila minima para que getDatosCDCASUB devuelva el row, pero con
    ' firmaAprobacionRespIngenieriaNombre vacia y solo Calidad firmada.
    ' EsAprobacionSuministradorCompleta exige las DOS firmas.
    Dim sql As String
    sql = "INSERT INTO tbDatosCDCASUB (idDatosCDCASUB, idSolicitud, " & _
          "refSubSuministrador, suministradorPrincipalNombreDir, " & _
          "subSuministradorNombreDir, requiereModificacionContrato, " & _
          "firmaAprobacionRespIngenieriaNombre, firmaAprobacionRespCalidadNombre) VALUES (" & _
          p_IdSolicitud & ", " & p_IdSolicitud & ", " & _
          "'REF-SUB-003', 'principal dir test', " & _
          "'sub dir test', False, " & _
          "'', 'Maria Calidad')"
    p_Db.Execute sql, dbFailOnError
End Sub

Private Sub SeedDatosCDCASUBRACCompleto(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long)
    ' Fila minima con racCodigo + racNombre poblados para que
    ' EsDictamenRACCompleto retorne True.
    Dim sql As String
    sql = "INSERT INTO tbDatosCDCASUB (idDatosCDCASUB, idSolicitud, " & _
          "racCodigo, racNombre) VALUES (" & _
          p_IdSolicitud & ", " & p_IdSolicitud & ", " & _
          "'RAC-2026-001', 'Juan RAC')"
    p_Db.Execute sql, dbFailOnError
End Sub

Private Sub SeedDatosCDCASUBDecisionFinalCompleta(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long)
    ' Fila minima con decisionFinal poblado para que
    ' EsDecisionFinalCompleta retorne True.
    Dim sql As String
    sql = "INSERT INTO tbDatosCDCASUB (idDatosCDCASUB, idSolicitud, " & _
          "decisionFinal) VALUES (" & _
          p_IdSolicitud & ", " & p_IdSolicitud & ", " & _
          "'APROBADO')"
    p_Db.Execute sql, dbFailOnError
End Sub

Private Sub SeedDatosCDCASUBDatosGeneralesCompleta(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long)
    ' Fila con los 4 campos obligatorios de EsDatosGeneralesCompleta poblados.
    Dim sql As String
    sql = "INSERT INTO tbDatosCDCASUB (idDatosCDCASUB, idSolicitud, " & _
          "refSubSuministrador, suministradorPrincipalNombreDir, " & _
          "subSuministradorNombreDir, requiereModificacionContrato) VALUES (" & _
          p_IdSolicitud & ", " & p_IdSolicitud & ", " & _
          "'REF-SUB-004', 'principal dir test', " & _
          "'sub dir test', False)"
    p_Db.Execute sql, dbFailOnError
End Sub

' CountDatosCDCASUB: explicit count against the injected db.
'   Replaces DCount, which always resolves through CurrentDb() (the frontend)
'   and therefore cannot see backend-only tables in the split architecture.
Private Function CountDatosCDCASUB(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long) As Long
    Dim rs As DAO.Recordset
    On Error GoTo EH
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS n FROM tbDatosCDCASUB WHERE idSolicitud=" & p_IdSolicitud, dbOpenSnapshot)
    If Not rs.EOF Then CountDatosCDCASUB = Nz(rs!n, 0)
    rs.Close
    Set rs = Nothing
    Exit Function
EH:
    CountDatosCDCASUB = -1
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
'   session cache pointing at the wrong db (TDD v2.4 §5.4).
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

Public Function Test_CDCASUB_Strict_EsParteTecnicaCompleta_TrueForCompleteRow() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New DatosCDCASUBServicio
    Dim countBefore As Long
    Dim result As Boolean

    logs = TestHelper.NewLogsArray(5)

    ' 1. Arrange — BeginTestSession + open sandbox + apply prod env
    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_CDCASUB_Strict_EsParteTecnicaCompleta_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Set db = OpenSandboxForTest()
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap005_cdcasub", "CAP005 Test User")
    logs(0) = "1. BeginTestSession OK; sandbox opened; prod env applied"

    ' 2. Fixture-first
    Call TeardownFixtures(db)
    Call SeedSolicitudCdcasub(db, ID_SOL_CDCASUB_PARTE, "CAP005-PARTE")
    Call SeedDatosCDCASUBParteTecnicaCompleta(db, ID_SOL_CDCASUB_PARTE)
    logs(1) = "2. Seeded solicitud CDCA-SUB with complete tbDatosCDCASUB (clasifNC=MAYOR, esSubSuministradorAD=True)"

    ' 3. Cardinality (TDD v2.4 §4.5)
    countBefore = CountDatosCDCASUB(db, ID_SOL_CDCASUB_PARTE)
    If countBefore <> 1 Then
        logs(2) = "3. FAILED: seed cardinalidad wrong: " & countBefore
        Test_CDCASUB_Strict_EsParteTecnicaCompleta_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("seed cardinalidad wrong", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. seed cardinalidad=" & countBefore

    ' 4. Act (read-only service method)
    result = svc.EsParteTecnicaCompleta(ID_SOL_CDCASUB_PARTE, db)
    logs(3) = "4. EsParteTecnicaCompleta returned: " & result

    ' 5. Assert
    If result <> True Then
        logs(4) = "5. FAILED: expected True (exclusive CDCASUB rule), got False"
        Test_CDCASUB_Strict_EsParteTecnicaCompleta_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("expected True", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. EsParteTecnicaCompleta=True as expected (CDCASUB-exclusive rule)"

    Test_CDCASUB_Strict_EsParteTecnicaCompleta_TrueForCompleteRow = _
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
    Test_CDCASUB_Strict_EsParteTecnicaCompleta_TrueForCompleteRow = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_CDCASUB_Strict_EsParteTecnicaCompleta_FalseWhenClasificacionNCInvalid() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New DatosCDCASUBServicio
    Dim result As Boolean

    logs = TestHelper.NewLogsArray(4)

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_CDCASUB_Strict_EsParteTecnicaCompleta_FalseWhenClasificacionNCInvalid = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Set db = OpenSandboxForTest()
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap005_cdcasub", "CAP005 Test User")
    logs(0) = "1. BeginTestSession OK; sandbox opened; prod env applied"

    Call TeardownFixtures(db)
    Call SeedSolicitudCdcasub(db, ID_SOL_CDCASUB_SAD_CLASIF, "CAP005-SAD-CLASIF")
    Call SeedDatosCDCASUBParteTecnicaClasificacionInvalida(db, ID_SOL_CDCASUB_SAD_CLASIF)
    logs(1) = "2. Seeded solicitud CDCA-SUB with clasificacionNC='INVALID' (not MAYOR/MENOR)"

    result = svc.EsParteTecnicaCompleta(ID_SOL_CDCASUB_SAD_CLASIF, db)
    logs(2) = "3. EsParteTecnicaCompleta returned: " & result

    If result <> False Then
        logs(3) = "4. FAILED: expected False (clasificacionNC not in {MAYOR, MENOR}), got True"
        Test_CDCASUB_Strict_EsParteTecnicaCompleta_FalseWhenClasificacionNCInvalid = _
            TestHelper.BuildJsonFail("expected False", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. EsParteTecnicaCompleta=False as expected (clasificacionNC invalid)"

    Test_CDCASUB_Strict_EsParteTecnicaCompleta_FalseWhenClasificacionNCInvalid = _
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
    Test_CDCASUB_Strict_EsParteTecnicaCompleta_FalseWhenClasificacionNCInvalid = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_CDCASUB_Strict_EsAprobacionSuministradorCompleta_FalseWhenIngenieriaEmpty() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New DatosCDCASUBServicio
    Dim result As Boolean

    logs = TestHelper.NewLogsArray(4)

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_CDCASUB_Strict_EsAprobacionSuministradorCompleta_FalseWhenIngenieriaEmpty = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Set db = OpenSandboxForTest()
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap005_cdcasub", "CAP005 Test User")
    logs(0) = "1. BeginTestSession OK; sandbox opened; prod env applied"

    Call TeardownFixtures(db)
    Call SeedSolicitudCdcasub(db, ID_SOL_CDCASUB_SAD_FIRMA, "CAP005-SAD-FIRMA")
    Call SeedDatosCDCASUBAprobacionSuministradorSinIngenieria(db, ID_SOL_CDCASUB_SAD_FIRMA)
    logs(1) = "2. Seeded solicitud CDCA-SUB with firmaAprobacionRespIngenieriaNombre='' and Calidad firma='Maria Calidad'"

    result = svc.EsAprobacionSuministradorCompleta(ID_SOL_CDCASUB_SAD_FIRMA, db)
    logs(2) = "3. EsAprobacionSuministradorCompleta returned: " & result

    If result <> False Then
        logs(3) = "4. FAILED: expected False (Ingenieria nombre vacio), got True"
        Test_CDCASUB_Strict_EsAprobacionSuministradorCompleta_FalseWhenIngenieriaEmpty = _
            TestHelper.BuildJsonFail("expected False", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. EsAprobacionSuministradorCompleta=False as expected (Ingenieria firma vacia)"

    Test_CDCASUB_Strict_EsAprobacionSuministradorCompleta_FalseWhenIngenieriaEmpty = _
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
    Test_CDCASUB_Strict_EsAprobacionSuministradorCompleta_FalseWhenIngenieriaEmpty = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_CDCASUB_Strict_EsDictamenRACCompleta_TrueForCompleteRow() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New DatosCDCASUBServicio
    Dim result As Boolean

    logs = TestHelper.NewLogsArray(4)

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_CDCASUB_Strict_EsDictamenRACCompleta_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Set db = OpenSandboxForTest()
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap005_cdcasub", "CAP005 Test User")
    logs(0) = "1. BeginTestSession OK; sandbox opened; prod env applied"

    Call TeardownFixtures(db)
    Call SeedSolicitudCdcasub(db, ID_SOL_CDCASUB_RAC, "CAP005-RAC")
    Call SeedDatosCDCASUBRACCompleto(db, ID_SOL_CDCASUB_RAC)
    logs(1) = "2. Seeded solicitud CDCA-SUB with racCodigo='RAC-2026-001' and racNombre='Juan RAC'"

    result = svc.EsDictamenRACCompleto(ID_SOL_CDCASUB_RAC, db)
    logs(2) = "3. EsDictamenRACCompleto returned: " & result

    If result <> True Then
        logs(3) = "4. FAILED: expected True (racCodigo + racNombre), got " & result
        Test_CDCASUB_Strict_EsDictamenRACCompleta_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("expected True", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. EsDictamenRACCompleto=True as expected"

    Test_CDCASUB_Strict_EsDictamenRACCompleta_TrueForCompleteRow = _
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
    Test_CDCASUB_Strict_EsDictamenRACCompleta_TrueForCompleteRow = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_CDCASUB_Strict_EsDecisionFinalCompleta_TrueForCompleteRow() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New DatosCDCASUBServicio
    Dim result As Boolean

    logs = TestHelper.NewLogsArray(4)

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_CDCASUB_Strict_EsDecisionFinalCompleta_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Set db = OpenSandboxForTest()
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap005_cdcasub", "CAP005 Test User")
    logs(0) = "1. BeginTestSession OK; sandbox opened; prod env applied"

    Call TeardownFixtures(db)
    Call SeedSolicitudCdcasub(db, ID_SOL_CDCASUB_DECISION, "CAP005-DECISION")
    Call SeedDatosCDCASUBDecisionFinalCompleta(db, ID_SOL_CDCASUB_DECISION)
    logs(1) = "2. Seeded solicitud CDCA-SUB with decisionFinal='APROBADO'"

    result = svc.EsDecisionFinalCompleta(ID_SOL_CDCASUB_DECISION, db)
    logs(2) = "3. EsDecisionFinalCompleta returned: " & result

    If result <> True Then
        logs(3) = "4. FAILED: expected True (decisionFinal poblada), got " & result
        Test_CDCASUB_Strict_EsDecisionFinalCompleta_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("expected True", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. EsDecisionFinalCompleta=True as expected"

    Test_CDCASUB_Strict_EsDecisionFinalCompleta_TrueForCompleteRow = _
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
    Test_CDCASUB_Strict_EsDecisionFinalCompleta_TrueForCompleteRow = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_CDCASUB_Strict_EsDatosGeneralesCompleta_TrueForCompleteRow() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New DatosCDCASUBServicio
    Dim result As Boolean

    logs = TestHelper.NewLogsArray(4)

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_CDCASUB_Strict_EsDatosGeneralesCompleta_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Set db = OpenSandboxForTest()
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap005_cdcasub", "CAP005 Test User")
    logs(0) = "1. BeginTestSession OK; sandbox opened; prod env applied"

    Call TeardownFixtures(db)
    Call SeedSolicitudCdcasub(db, ID_SOL_CDCASUB_GENERALES, "CAP005-GEN")
    Call SeedDatosCDCASUBDatosGeneralesCompleta(db, ID_SOL_CDCASUB_GENERALES)
    logs(1) = "2. Seeded solicitud CDCA-SUB with 4 datos generales obligatorios"

    result = svc.EsDatosGeneralesCompleta(ID_SOL_CDCASUB_GENERALES, db)
    logs(2) = "3. EsDatosGeneralesCompleta returned: " & result

    If result <> True Then
        logs(3) = "4. FAILED: expected True (4 campos), got " & result
        Test_CDCASUB_Strict_EsDatosGeneralesCompleta_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("expected True", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. EsDatosGeneralesCompleta=True as expected"

    Test_CDCASUB_Strict_EsDatosGeneralesCompleta_TrueForCompleteRow = _
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
    Test_CDCASUB_Strict_EsDatosGeneralesCompleta_TrueForCompleteRow = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ----------------------------------------------------------------------------
' Smoke aggregator — globally unique Public Function per §1.1.1
' ----------------------------------------------------------------------------
Public Function Test_CDCASUB_Strict_RunAll() As String
    On Error GoTo EH
    Dim logs() As String
    Dim result As String

    logs = TestHelper.NewLogsArray(7)

    result = Test_CDCASUB_Strict_EsParteTecnicaCompleta_TrueForCompleteRow()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(0) = "1. EsParteTecnicaCompleta=True on complete CDCASUB row (exclusive rule)"

    result = Test_CDCASUB_Strict_EsParteTecnicaCompleta_FalseWhenClasificacionNCInvalid()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(1) = "2. EsParteTecnicaCompleta=False when clasificacionNC='INVALID' (sad path)"

    result = Test_CDCASUB_Strict_EsAprobacionSuministradorCompleta_FalseWhenIngenieriaEmpty()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(2) = "3. EsAprobacionSuministradorCompleta=False when Ingenieria firma vacia (sad path)"

    result = Test_CDCASUB_Strict_EsDictamenRACCompleta_TrueForCompleteRow()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(3) = "4. EsDictamenRACCompleto=True on RAC canonico"

    result = Test_CDCASUB_Strict_EsDecisionFinalCompleta_TrueForCompleteRow()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(4) = "5. EsDecisionFinalCompleta=True on decisionFinal poblada"

    result = Test_CDCASUB_Strict_EsDatosGeneralesCompleta_TrueForCompleteRow()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(5) = "6. EsDatosGeneralesCompleta=True on 4 datos generales"

    logs(6) = "7. PASS"
    Test_CDCASUB_Strict_RunAll = TestHelper.BuildJsonOk("true", logs)
    Exit Function

Failed:
    Test_CDCASUB_Strict_RunAll = TestHelper.BuildJsonFail("one CAP-005 CDCASUB atom failed", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "Test_CDCASUB_Strict_RunAll: ERR " & Err.Number & " - " & Err.Description
    Test_CDCASUB_Strict_RunAll = TestHelper.BuildJsonFail(Err.Description, logs)
End Function
