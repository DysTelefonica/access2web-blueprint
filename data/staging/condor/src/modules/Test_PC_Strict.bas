Attribute VB_Name = "Test_PC_Strict"
Option Compare Database
Option Explicit

' ============================================================================
' Test_PC_Strict — CAP-004 PC strict TDD v2.4.2 atoms (Fase B1, 2026-06-15)
'
' Touched tables: tbSolicitudes, TbExpedientes, tbDatosPC.
' Read + write via explicit DAO.Database injection. Cardinalidad on
' mutations. Fixture IDs in 900710-900799.
'
' Note: EsSolicitudEnValidacion is private and reads via getdb() when db is
' not passed. For the sandbox path this resolves correctly because m_TestingMode
' is set; for unit testing with mocks, that helper would need to be exposed.
' We test the Public API only here.
'
' v2.4.2 migration: SetupPcSandbox (private) replaced with the canonical
'   TestHelper.BeginTestSession(logs, errMsg).
'   TestHelper.GetTestDb() replaced with explicit OpenSandboxForTest()
'   helper that calls DBEngine.Workspaces(0).OpenDatabase(
'   m_BackendSandboxURL, ..., ";PWD=" & m_BackendSandboxPassword).
'   The GuardarAprobacionSuministrador atom re-opens dbPost via getdb()
'   after the call, because the service's cleanup does Set db = Nothing
'   (a separate pre-fix bug tracked outside this migration).
' v2.4.3 consolidation: ApplyProdEnvForTest / EnsureFolder / CleanupTempRoot
'   (private) replaced with TestHelper.SetupProdGlobalsForTest /
'   TestHelper.CleanupProdTempRoot canónicos.
' ============================================================================

Private Const TEST_ID_BASE As Long = 900710
Private Const TEST_ID_TOP As Long = 900799

Private Const ID_SOL_PC_DATOS As Long = 900711
Private Const ID_SOL_PC_RAC As Long = 900712
Private Const ID_SOL_PC_PROPUESTA As Long = 900713
Private Const ID_SOL_PC_GENERALES As Long = 900714
Private Const ID_SOL_PC_PARTE_TECNICA As Long = 900715
Private Const ID_SOL_PC_MOTIVOS As Long = 900716
Private Const ID_SOL_PC_INVAL As Long = 900717
Private Const ID_SOL_PC_NOINVAL As Long = 900718

Private Const FILE_PREFIX As String = "CAP004-STRICT-"

' ----------------------------------------------------------------------------
' Fixture helpers (schema-first inserts against the real ERD)
' ----------------------------------------------------------------------------

Private Sub TeardownFixtures(ByVal p_Db As DAO.Database)
    On Error Resume Next
    p_Db.Execute "DELETE FROM tbDatosPC WHERE idSolicitud >= " & TEST_ID_BASE & " AND idSolicitud <= " & TEST_ID_TOP, dbFailOnError
    p_Db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud >= " & TEST_ID_BASE & " AND idSolicitud <= " & TEST_ID_TOP, dbFailOnError
    p_Db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente >= " & TEST_ID_BASE & " AND IDExpediente <= " & TEST_ID_TOP, dbFailOnError
    On Error GoTo 0
End Sub

Private Sub SeedSolicitud(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long, ByVal p_Codigo As String)
    On Error Resume Next
    p_Db.Execute "DELETE FROM tbDatosPC WHERE idSolicitud=" & p_IdSolicitud, dbFailOnError
    p_Db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud=" & p_IdSolicitud, dbFailOnError
    p_Db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & p_IdSolicitud, dbFailOnError
    On Error GoTo 0
    p_Db.Execute "INSERT INTO TbExpedientes (IDExpediente) VALUES (" & p_IdSolicitud & ")", dbFailOnError
    Dim sql As String
    sql = "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, " & _
          "idEstadoInterno, fechaCreacion, usuarioCreacion, revisionCalidadEstado) VALUES (" & _
          p_IdSolicitud & ", " & p_IdSolicitud & ", 'PC', '" & p_Codigo & "', 2, #2025-01-01 10:00:00#, " & _
          "'TestPcStrict', 'PENDIENTE')"
    p_Db.Execute sql, dbFailOnError
End Sub

Private Sub SeedDatosPCCompleto(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long)
    ' Seed a tbDatosPC row with the EXACT column names from ERD condor_datos.md §tbDatosPC.
    ' Used by EsDetalleCompleto (requires descripcionMaterialAfectado and
    ' descripcionPropuestaCambio both populated) and as the base row for
    ' GuardarAprobacionSuministrador (an UPDATE).
    Dim sql As String
    sql = "INSERT INTO tbDatosPC (idDatosPC, idSolicitud, " & _
          "refContratoInspeccionOficial, refSuministrador, " & _
          "denominacionContrato, suministradorNombreDir, objetoContrato, " & _
          "descripcionMaterialAfectado, numPlanoEspecificacion, " & _
          "descripcionPropuestaCambio, descripcionPropuestaCambioCont, " & _
          "motivoCorregirDeficiencias, motivoMejorarCapacidad, " & _
          "motivoAumentarNacionalizacion, motivoMejorarSeguridad, " & _
          "motivoMejorarFiabilidad, motivoMejorarCosteEficacia, " & _
          "motivoOtros, motivoOtrosDetalle, " & _
          "incidenciaCoste, incidenciaPlazo, " & _
          "incidenciaSeguridad, incidenciaFiabilidad, " & _
          "incidenciaMantenibilidad, incidenciaIntercambiabilidad, " & _
          "incidenciaVidaUtilAlmacen, incidenciaFuncionamientoFuncion, " & _
          "impactoClasificacion, CambioAfectaAMaterial) VALUES (" & _
          p_IdSolicitud & ", " & _
          "'REF-TEST-001', 'SUM-TEST-001', " & _
          "'denominacion test', 'suministrador test', 'objeto test', " & _
          "'material de prueba', 'plano-001', " & _
          "'propuesta de cambio test', 'continuacion propuesta', " & _
          "True, False, False, False, False, False, " & _
          "False, 'otros motivo detalle', " & _
          "'incidencia coste', 'incidencia plazo', " & _
          "True, False, False, False, False, False, " & _
          "'clasificacion impacto', 'S')"
    sql = Replace(sql, "VALUES (" & p_IdSolicitud & ",", "VALUES (" & p_IdSolicitud & ", " & p_IdSolicitud & ",")
    p_Db.Execute sql, dbFailOnError
End Sub

Private Sub SeedDatosPCParteTecnicaCompleta(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long)
    ' Seed a tbDatosPC row with the CANONICAL values required by EsParteTecnicaCompleta
    ' and EsMotivosCompleto: incidenciaCoste in {AUMENTARA, DISMINUIRA, NO VARIARA}
    ' and impactoClasificacion in {MAYOR, MENOR}, with al menos un motivo=True.
    Dim sql As String
    sql = "INSERT INTO tbDatosPC (idDatosPC, idSolicitud, " & _
          "refContratoInspeccionOficial, refSuministrador, " & _
          "denominacionContrato, suministradorNombreDir, objetoContrato, " & _
          "descripcionMaterialAfectado, descripcionPropuestaCambio, " & _
          "motivoCorregirDeficiencias, motivoMejorarCapacidad, " & _
          "motivoAumentarNacionalizacion, motivoMejorarSeguridad, " & _
          "motivoMejorarFiabilidad, motivoMejorarCosteEficacia, " & _
          "motivoOtros, " & _
          "incidenciaCoste, " & _
          "impactoClasificacion, CambioAfectaAMaterial) VALUES (" & _
          p_IdSolicitud & ", " & p_IdSolicitud & ", " & _
          "'REF-TEST-001', 'SUM-TEST-001', " & _
          "'denominacion test', 'suministrador test', 'objeto test', " & _
          "'material de prueba', 'propuesta de cambio test', " & _
          "True, False, False, False, False, False, " & _
          "False, " & _
          "'AUMENTARÁ', " & _
          "'MAYOR', 'S')"
    p_Db.Execute sql, dbFailOnError
End Sub

' CountDatosPC: explicit count against the injected db.
Private Function CountDatosPC(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long) As Long
    Dim rs As DAO.Recordset
    On Error GoTo EH
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS n FROM tbDatosPC WHERE idSolicitud=" & p_IdSolicitud, dbOpenSnapshot)
    If Not rs.EOF Then CountDatosPC = Nz(rs!n, 0)
    rs.Close
    Set rs = Nothing
    Exit Function
EH:
    CountDatosPC = -1
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
End Function

Private Function GetFirmaOficinaTecnica(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long) As String
    Dim rs As DAO.Recordset
    On Error GoTo EH
    Set rs = p_Db.OpenRecordset("SELECT firmaOficinaTecnicaNombre FROM tbDatosPC WHERE idSolicitud=" & p_IdSolicitud, dbOpenSnapshot)
    If Not rs.EOF Then GetFirmaOficinaTecnica = Nz(rs!firmaOficinaTecnicaNombre, "")
    rs.Close
    Set rs = Nothing
    Exit Function
EH:
    GetFirmaOficinaTecnica = ""
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
End Function

Private Function GetFirmaRepSuministrador(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long) As String
    Dim rs As DAO.Recordset
    On Error GoTo EH
    Set rs = p_Db.OpenRecordset("SELECT firmaRepSuministradorNombre FROM tbDatosPC WHERE idSolicitud=" & p_IdSolicitud, dbOpenSnapshot)
    If Not rs.EOF Then GetFirmaRepSuministrador = Nz(rs!firmaRepSuministradorNombre, "")
    rs.Close
    Set rs = Nothing
    Exit Function
EH:
    GetFirmaRepSuministrador = ""
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

Public Function Test_PC_Strict_EsDetalleCompleto_TrueForCompleteRow() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New DatosPCServicio

    logs = TestHelper.NewLogsArray(3)

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_PC_Strict_EsDetalleCompleto_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Set db = OpenSandboxForTest()
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap004_pc", "CAP004 Test User")
    logs(0) = "1. BeginTestSession OK; sandbox opened; prod env applied"

    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_PC_DATOS, "CAP004-DATOS")
    Call SeedDatosPCCompleto(db, ID_SOL_PC_DATOS)
    logs(1) = "2. Seeded solicitud with complete tbDatosPC row"

    If Not svc.EsDetalleCompleto(ID_SOL_PC_DATOS, db) Then
        logs(2) = "3. FAILED: EsDetalleCompleto=False on complete row"
        Test_PC_Strict_EsDetalleCompleto_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("EsDetalleCompleto=False", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. EsDetalleCompleto=True on complete row"

    Test_PC_Strict_EsDetalleCompleto_TrueForCompleteRow = _
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
    Test_PC_Strict_EsDetalleCompleto_TrueForCompleteRow = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PC_Strict_EsDetalleCompleto_FalseWhenDescripcionMaterialEmpty() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New DatosPCServicio

    logs = TestHelper.NewLogsArray(3)

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_PC_Strict_EsDetalleCompleto_FalseWhenDescripcionMaterialEmpty = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Set db = OpenSandboxForTest()
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap004_pc", "CAP004 Test User")
    logs(0) = "1. BeginTestSession OK; sandbox opened; prod env applied"

    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_PC_DATOS, "CAP004-DATOS")
    Call SeedDatosPCCompleto(db, ID_SOL_PC_DATOS)
    db.Execute "UPDATE tbDatosPC SET descripcionMaterialAfectado=NULL WHERE idSolicitud=" & ID_SOL_PC_DATOS, dbFailOnError
    logs(1) = "2. Seeded solicitud then set descripcionMaterialAfectado=NULL"

    If svc.EsDetalleCompleto(ID_SOL_PC_DATOS, db) Then
        logs(2) = "3. FAILED: EsDetalleCompleto=True with empty descripcionMaterialAfectado"
        Test_PC_Strict_EsDetalleCompleto_FalseWhenDescripcionMaterialEmpty = _
            TestHelper.BuildJsonFail("EsDetalleCompleto=True", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. EsDetalleCompleto=False when descripcionMaterialAfectado is NULL"

    Test_PC_Strict_EsDetalleCompleto_FalseWhenDescripcionMaterialEmpty = _
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
    Test_PC_Strict_EsDetalleCompleto_FalseWhenDescripcionMaterialEmpty = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PC_Strict_GuardarAprobacionSuministrador_PersistsFirmantes() As String
    ' Slice B1 (2026-06-15): GuardarAprobacionSuministrador is an UPDATE, not an
    ' INSERT. It requires a pre-existing tbDatosPC row (the 'Datos Generales'
    ' base). After the call, cardinalidad must be unchanged and the firmantes
    ' must be persisted.
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim dbPost As DAO.Database
    Dim tempRoot As String
    Dim svc As New DatosPCServicio
    Dim pc As New DatosPC
    Dim countBefore As Long
    Dim countAfter As Long
    Dim persistedFirmaOT As String
    Dim persistedFirmaRS As String

    logs = TestHelper.NewLogsArray(7)

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_PC_Strict_GuardarAprobacionSuministrador_PersistsFirmantes = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Set db = OpenSandboxForTest()
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap004_pc", "CAP004 Test User")
    logs(0) = "1. BeginTestSession OK; sandbox opened; prod env applied"

    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_PC_RAC, "CAP004-RAC")
    Call SeedDatosPCCompleto(db, ID_SOL_PC_RAC)
    logs(1) = "2. Seeded parent graph + DatosPC base row"

    countBefore = CountDatosPC(db, ID_SOL_PC_RAC)
    logs(2) = "3. countBefore=" & countBefore & " (after seed, before act)"

    pc.idSolicitud = ID_SOL_PC_RAC
    pc.firmaOficinaTecnicaNombre = "oficial-test"
    pc.firmaRepSuministradorNombre = "rep-test"
    logs(3) = "4. Built DatosPC with new firmantes (UPDATE payload)"

    svc.GuardarAprobacionSuministrador pc, db
    logs(4) = "5. GuardarAprobacionSuministrador called with explicit DAO.Database"

    ' NOTE: DatosPCRepositorio.ActualizarAprobacionSuministrador does
    ' `Set db = Nothing` in its cleanup, which destroys the caller's
    ' connection. We re-open a fresh DAO connection for the post-act
    ' assertions (tracked as a separate pre-fix bug outside this
    ' migration scope).
    Set dbPost = OpenSandboxForTest()

    countAfter = CountDatosPC(dbPost, ID_SOL_PC_RAC)
    logs(5) = "6. countAfter=" & countAfter & " (via fresh dbPost connection)"

    If countAfter <> countBefore Then
        Test_PC_Strict_GuardarAprobacionSuministrador_PersistsFirmantes = _
            TestHelper.BuildJsonFail("UPDATE changed cardinalidad: before=" & countBefore & " after=" & countAfter, logs)
        GoTo CleanExit
    End If
    logs(5) = logs(5) & " (unchanged, UPDATE confirmed)"

    persistedFirmaOT = GetFirmaOficinaTecnica(dbPost, ID_SOL_PC_RAC)
    persistedFirmaRS = GetFirmaRepSuministrador(dbPost, ID_SOL_PC_RAC)
    logs(6) = "7. persisted: OT=[" & persistedFirmaOT & "] RS=[" & persistedFirmaRS & "]"

    If persistedFirmaOT <> "oficial-test" Then
        Test_PC_Strict_GuardarAprobacionSuministrador_PersistsFirmantes = _
            TestHelper.BuildJsonFail("firmaOficinaTecnicaNombre not persisted: [" & persistedFirmaOT & "]", logs)
        GoTo CleanExit
    End If
    If persistedFirmaRS <> "rep-test" Then
        Test_PC_Strict_GuardarAprobacionSuministrador_PersistsFirmantes = _
            TestHelper.BuildJsonFail("firmaRepSuministradorNombre not persisted: [" & persistedFirmaRS & "]", logs)
        GoTo CleanExit
    End If

    logs(6) = "7. PASS"
    Test_PC_Strict_GuardarAprobacionSuministrador_PersistsFirmantes = _
        TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then
        TeardownFixtures db
        db.Close
        Set db = Nothing
    End If
    If Not dbPost Is Nothing Then
        dbPost.Close
        Set dbPost = Nothing
    End If
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(6) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_PC_Strict_GuardarAprobacionSuministrador_PersistsFirmantes = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PC_Strict_EsDatosGeneralesCompleta_TrueForCompleteRow() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New DatosPCServicio
    Dim result As Boolean

    logs = TestHelper.NewLogsArray(4)

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_PC_Strict_EsDatosGeneralesCompleta_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Set db = OpenSandboxForTest()
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap004_pc", "CAP004 Test User")
    logs(0) = "1. BeginTestSession OK; sandbox opened; prod env applied"

    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_PC_GENERALES, "CAP004-GEN")
    Call SeedDatosPCCompleto(db, ID_SOL_PC_GENERALES)
    logs(1) = "2. Seeded solicitud with complete tbDatosPC (5 campos datos generales)"

    result = svc.EsDatosGeneralesCompleta(ID_SOL_PC_GENERALES, db)
    logs(2) = "3. EsDatosGeneralesCompleta returned: " & result

    If result <> True Then
        logs(3) = "4. FAILED: expected True (5 campos datos generales), got " & result
        Test_PC_Strict_EsDatosGeneralesCompleta_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("expected True", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. EsDatosGeneralesCompleta=True as expected"

    Test_PC_Strict_EsDatosGeneralesCompleta_TrueForCompleteRow = _
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
    Test_PC_Strict_EsDatosGeneralesCompleta_TrueForCompleteRow = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PC_Strict_EsParteTecnicaCompleta_TrueForCompleteRow() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New DatosPCServicio
    Dim result As Boolean

    logs = TestHelper.NewLogsArray(4)

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_PC_Strict_EsParteTecnicaCompleta_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Set db = OpenSandboxForTest()
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap004_pc", "CAP004 Test User")
    logs(0) = "1. BeginTestSession OK; sandbox opened; prod env applied"

    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_PC_PARTE_TECNICA, "CAP004-PARTE")
    Call SeedDatosPCParteTecnicaCompleta(db, ID_SOL_PC_PARTE_TECNICA)
    logs(1) = "2. Seeded solicitud with incidenciaCoste='AUMENTARÁ', impactoClasificacion='MAYOR', motivoCorregirDeficiencias=True"

    result = svc.EsParteTecnicaCompleta(ID_SOL_PC_PARTE_TECNICA, db)
    logs(2) = "3. EsParteTecnicaCompleta returned: " & result

    If result <> True Then
        logs(3) = "4. FAILED: expected True (parte tecnica canónica), got " & result
        Test_PC_Strict_EsParteTecnicaCompleta_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("expected True", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. EsParteTecnicaCompleta=True as expected"

    Test_PC_Strict_EsParteTecnicaCompleta_TrueForCompleteRow = _
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
    Test_PC_Strict_EsParteTecnicaCompleta_TrueForCompleteRow = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PC_Strict_EsMotivosCompleto_TrueForCompleteRow() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New DatosPCServicio
    Dim result As Boolean

    logs = TestHelper.NewLogsArray(4)

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_PC_Strict_EsMotivosCompleto_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Set db = OpenSandboxForTest()
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap004_pc", "CAP004 Test User")
    logs(0) = "1. BeginTestSession OK; sandbox opened; prod env applied"

    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_PC_MOTIVOS, "CAP004-MOTIVOS")
    Call SeedDatosPCParteTecnicaCompleta(db, ID_SOL_PC_MOTIVOS)
    logs(1) = "2. Seeded solicitud with motivos+clasificacion canonicos"

    result = svc.EsMotivosCompleto(ID_SOL_PC_MOTIVOS, db)
    logs(2) = "3. EsMotivosCompleto returned: " & result

    If result <> True Then
        logs(3) = "4. FAILED: expected True (motivos canónicos), got " & result
        Test_PC_Strict_EsMotivosCompleto_TrueForCompleteRow = _
            TestHelper.BuildJsonFail("expected True", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. EsMotivosCompleto=True as expected"

    Test_PC_Strict_EsMotivosCompleto_TrueForCompleteRow = _
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
    Test_PC_Strict_EsMotivosCompleto_TrueForCompleteRow = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PC_Strict_EsSolicitudEnValidacion_TrueAndFalse() As String
    On Error GoTo EH
    Dim logs() As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New DatosPCServicio
    Dim resultInVal As Boolean
    Dim resultNotInVal As Boolean

    logs = TestHelper.NewLogsArray(6)

    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(0) = "1. TESTS BLOCKED: " & errMsg
        Test_PC_Strict_EsSolicitudEnValidacion_TrueAndFalse = _
            TestHelper.BuildJsonFail("TESTS BLOCKED: " & errMsg, logs)
        GoTo CleanExit
    End If
    Set db = OpenSandboxForTest()
    Call TestHelper.SetupProdGlobalsForTest(errMsg, tempRoot, "cap004_pc", "CAP004 Test User")
    logs(0) = "1. BeginTestSession OK; sandbox opened; prod env applied"

    Call TeardownFixtures(db)

    ' Seed solicitud en estadoValidacion (5)
    Call SeedSolicitud(db, ID_SOL_PC_INVAL, "CAP004-INVAL")
    db.Execute "UPDATE tbSolicitudes SET idEstadoInterno=5 WHERE idSolicitud=" & ID_SOL_PC_INVAL, dbFailOnError
    logs(1) = "2. Seeded solicitud id=" & ID_SOL_PC_INVAL & " with idEstadoInterno=5 (estadoValidacion)"

    ' Seed solicitud en estadoRegistro (2)
    Call SeedSolicitud(db, ID_SOL_PC_NOINVAL, "CAP004-NOINVAL")
    logs(2) = "3. Seeded solicitud id=" & ID_SOL_PC_NOINVAL & " with idEstadoInterno=2 (estadoRegistro, default)"

    resultInVal = svc.EsSolicitudEnValidacion(ID_SOL_PC_INVAL, db)
    logs(3) = "4. EsSolicitudEnValidacion(estadoValidacion) returned: " & resultInVal

    If resultInVal <> True Then
        logs(4) = "5. FAILED: expected True for estadoValidacion (5), got " & resultInVal
        Test_PC_Strict_EsSolicitudEnValidacion_TrueAndFalse = _
            TestHelper.BuildJsonFail("expected True for estadoValidacion", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. resultInVal=True as expected"

    resultNotInVal = svc.EsSolicitudEnValidacion(ID_SOL_PC_NOINVAL, db)
    logs(5) = "6. EsSolicitudEnValidacion(estadoRegistro) returned: " & resultNotInVal

    If resultNotInVal <> False Then
        logs(5) = "6. FAILED: expected False for estadoRegistro (2), got " & resultNotInVal
        Test_PC_Strict_EsSolicitudEnValidacion_TrueAndFalse = _
            TestHelper.BuildJsonFail("expected False for estadoRegistro", logs)
        GoTo CleanExit
    End If
    logs(5) = "6. resultNotInVal=False as expected"

    Test_PC_Strict_EsSolicitudEnValidacion_TrueAndFalse = _
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
    Test_PC_Strict_EsSolicitudEnValidacion_TrueAndFalse = _
        TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ----------------------------------------------------------------------------
' Smoke aggregator — globally unique Public Function per §1.1.1
' ----------------------------------------------------------------------------
Public Function Test_PC_Strict_RunAll() As String
    On Error GoTo EH
    Dim logs() As String
    Dim result As String

    logs = TestHelper.NewLogsArray(8)

    result = Test_PC_Strict_EsDetalleCompleto_TrueForCompleteRow()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(0) = "1. EsDetalleCompleto=True on complete row"

    result = Test_PC_Strict_EsDetalleCompleto_FalseWhenDescripcionMaterialEmpty()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(1) = "2. EsDetalleCompleto=False when descripcionMaterialAfectado is NULL"

    result = Test_PC_Strict_EsDatosGeneralesCompleta_TrueForCompleteRow()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(2) = "3. EsDatosGeneralesCompleta=True on complete row (5 campos datos generales)"

    result = Test_PC_Strict_EsParteTecnicaCompleta_TrueForCompleteRow()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(3) = "4. EsParteTecnicaCompleta=True with canonical values"

    result = Test_PC_Strict_EsMotivosCompleto_TrueForCompleteRow()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(4) = "5. EsMotivosCompleto=True with motivos canonicos"

    result = Test_PC_Strict_GuardarAprobacionSuministrador_PersistsFirmantes()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(5) = "6. GuardarAprobacionSuministrador persists firmantes"

    result = Test_PC_Strict_EsSolicitudEnValidacion_TrueAndFalse()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(6) = "7. EsSolicitudEnValidacion True/False for estadoValidacion/Registro"

    logs(7) = "8. PASS"
    Test_PC_Strict_RunAll = TestHelper.BuildJsonOk("true", logs)
    Exit Function

Failed:
    Test_PC_Strict_RunAll = TestHelper.BuildJsonFail("one CAP-004 PC atom failed", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PC_Strict_RunAll: ERR " & Err.Number & " - " & Err.Description
    Test_PC_Strict_RunAll = TestHelper.BuildJsonFail(Err.Description, logs)
End Function
