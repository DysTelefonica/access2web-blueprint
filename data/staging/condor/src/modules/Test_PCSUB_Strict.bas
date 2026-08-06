Attribute VB_Name = "Test_PCSUB_Strict"
Option Compare Database
Option Explicit

' ============================================================================
' Test_PCSUB_Strict — v2.4.2 strict TDD atoms for CAP-001 (PCSUB)
'
' Skill:    access-vba-tdd v2.4.2
' SDD/CAP:  CAP-001 (PCSUB) — business rules BR-001..BR-011
' Branch:   staging
' Project:  condor (Dysflow projectId)
'
' Why a new file (instead of editing Test_PCSUB.bas)
'   Test_PCSUB.bas carries 2036 lines of legacy work: 15 strong atoms
'   (Slices C/D/E/F + Issue7) plus ~50 fake-stub groups (Groups 1-18)
'   that return BuildJsonOk("true", logs) without any real assertion.
'   Those fake stubs violate v2.4.2 §4.2 (no-humo). Refactoring that
'   file in a single slice would breach the 400-line review budget and
'   mix strict migration with legacy debt. The strict atoms live here,
'   in a self-contained file, modeled on Test_Issue7_DbInjection.bas
'   and Test_Issue19_EstadoValidacion.bas.
'
' What this file asserts (per BR-001..BR-011 of CAP-001 §2)
'   BR-001  DatosGenerales required fields        — 4 atoms
'   BR-002  Propuesta required fields             — 4 atoms
'   BR-003  Impacto motivos (>=1 motivo, otros)   — 3 atoms
'   BR-004  IncidenciaCoste/Plazo enum            — 2 atoms
'   BR-005  ImpactoClasificacion MAYOR/MENOR      — 1 atom
'   BR-006  CambioAfectaAMaterial enum            — 1 atom
'   BR-007  AprobacionSuministrador two signers   — 3 atoms
'   BR-008  DictamenRAC: non-RECHAZADO requires code — 3 atoms
'   BR-009  DecisionFinal required               — 2 atoms
'   BR-010  Limpiar* side-effects (clears only target block) — 4 atoms
'   BR-011  Injected DAO.Database preservation    — 2 atoms
'   Total: 29 atomic Public Function tests, one per manifest entry.
'
' TDD skill v2.4.2 compliance
'   - Public Function Test_*() As String returning JSON {ok,value,payload,error,logs}
'   - Schema-first seed (only Required/NotNull columns from docs/ERD/condor_datos.md)
'   - Fixture IDs deterministic >= 900100..900199 to avoid collision with Issue7
'     (900001) and Issue19 (900010/900011) and legacy Test_PCSUB.bas IDs
'   - DAO.Database injected explicitly (no Nothing as db)
'   - countBefore / countAfter for every mutation (TDD v2.4.2 §4.5)
'   - SetupSandbox validates sandbox path (UNC/production rejected) and
'     returns "TESTS BLOCKED" on failure without touching data
'   - TeardownAll deletes only deterministic fixture IDs (>= 900100)
'   - Cero Debug.Print, MsgBox, mutation of TbConfiguracionBackends
'   - No EVE() (reserved for suite lifecycle; the per-file SuiteSetup
'     legacy in Test_PCSUB.bas is replaced by inline SetupSandbox here)
'   - No mocks; the tests exercise the real service and repository
'     against the real sandbox db (TDD v2.4.2 §4.6)
' ============================================================================

' --- Fixture ID ranges (TDD v2.4.2 §5.1) ---
Private Const TEST_ID_BASE As Long = 900100
Private Const TEST_ID_TOP As Long = 900199

' --- BR-001 fixtures (4 atoms use 900110..900113) ---
Private Const BR001_ID_HAPPY As Long = 900110
Private Const BR001_ID_REJ_REF As Long = 900111
Private Const BR001_ID_COMP As Long = 900112
Private Const BR001_ID_INCOMP As Long = 900113

' --- BR-002 fixtures (4 atoms use 900120..900123) ---
Private Const BR002_ID_HAPPY As Long = 900120
Private Const BR002_ID_NO_MATERIAL As Long = 900121
Private Const BR002_ID_NO_PROPUESTA As Long = 900122
Private Const BR002_ID_COMP As Long = 900123

' --- BR-003 fixtures (3 atoms use 900130..900132) ---
Private Const BR003_ID_NO_MOTIVO As Long = 900130
Private Const BR003_ID_OTROS_DETALLE_NO_FLAG As Long = 900131
Private Const BR003_ID_COMP As Long = 900132

' --- BR-004 fixtures (2 atoms use 900140..900141) ---
Private Const BR004_ID_BAD_COSTE As Long = 900140
Private Const BR004_ID_BAD_PLAZO As Long = 900141

' --- BR-005 fixture (1 atom uses 900150) ---
Private Const BR005_ID_BAD_CLASIF As Long = 900150

' --- BR-006 fixture (1 atom uses 900160) ---
Private Const BR006_ID_BAD_MATERIAL As Long = 900160

' --- BR-007 fixtures (3 atoms use 900170..900172) ---
Private Const BR007_ID_HAPPY As Long = 900170
Private Const BR007_ID_REJ_S1 As Long = 900171
Private Const BR007_ID_REJ_S2 As Long = 900172

' --- BR-008 fixtures (3 atoms use 900180..900182) ---
Private Const BR008_ID_HAPPY As Long = 900180
Private Const BR008_ID_REJ_EMPTY As Long = 900181
Private Const BR008_ID_REJ_NONREJ As Long = 900182

' --- BR-009 fixtures (2 atoms use 900185..900186) ---
Private Const BR009_ID_HAPPY As Long = 900185
Private Const BR009_ID_EMPTY As Long = 900186

' --- BR-010 fixtures (4 atoms use 900190..900193) ---
Private Const BR010_ID_PURGA As Long = 900190
Private Const BR010_ID_DEL_RAC As Long = 900191
Private Const BR010_ID_DEL_APROB As Long = 900192
Private Const BR010_ID_DEL_DEC As Long = 900193

' --- BR-011 fixtures (2 atoms use 900195..900196) ---
Private Const BR011_ID_GUARDAR As Long = 900195
Private Const BR011_ID_ACTUALIZAR As Long = 900196

' ============================================================================
' SETUP / TEARDOWN — self-contained, mirrors Test_Issue7_DbInjection.bas
' ============================================================================

Private Sub SetupSandbox(ByRef p_Error As String)
    On Error GoTo EH
    Dim dbFrontend As DAO.Database
    Dim rs As DAO.Recordset
    Dim sandboxPath As String
    Dim backendPassword As String
    Dim fsoLocal As Object
    Dim dbProbe As DAO.Database

    Set dbFrontend = CurrentDb
    Set rs = dbFrontend.OpenRecordset( _
        "SELECT TOP 1 BackendSandbox, BackendTest, PasswordBackend " & _
        "FROM TbConfiguracionBackends WHERE Habilitado=True", dbOpenSnapshot)
    If rs.EOF Then
        p_Error = "TESTS BLOCKED: TbConfiguracionBackends has no enabled row"
        GoTo CleanExit
    End If
    sandboxPath = Trim$(Nz(rs!BackendSandbox, ""))
    If Len(sandboxPath) = 0 Then sandboxPath = Trim$(Nz(rs!BackendTest, ""))
    backendPassword = Nz(rs!PasswordBackend, GetPasswordDB())
    rs.Close: Set rs = Nothing
    Set dbFrontend = Nothing

    If Len(sandboxPath) = 0 Then
        p_Error = "TESTS BLOCKED: BackendSandbox/BackendTest path is empty"
        Exit Sub
    End If

    If Left$(sandboxPath, 2) = "\\" Then
        p_Error = "TESTS BLOCKED: sandbox backend must be local, not UNC: " & sandboxPath
        Exit Sub
    End If

    If InStr(1, LCase$(sandboxPath), "condor_datos.accdb", vbTextCompare) = 0 Then
        p_Error = "TESTS BLOCKED: sandbox backend path does not match CONDOR fingerprint: " & sandboxPath
        Exit Sub
    End If

    If InStr(1, LCase$(sandboxPath), "\\datoste\\", vbTextCompare) > 0 Then
        p_Error = "TESTS BLOCKED: sandbox backend path looks productive: " & sandboxPath
        Exit Sub
    End If

    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    If Not fsoLocal.FileExists(sandboxPath) Then
        Set fsoLocal = Nothing
        p_Error = "TESTS BLOCKED: sandbox backend not found: " & sandboxPath
        Exit Sub
    End If
    Set fsoLocal = Nothing

    Set dbProbe = DBEngine.Workspaces(0).OpenDatabase(sandboxPath, False, True, ";PWD=" & backendPassword)
    dbProbe.Close: Set dbProbe = Nothing

    m_BackendSandboxURL = sandboxPath
    m_PasswordBackend = backendPassword
    m_TestingMode = True
    Exit Sub

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set dbFrontend = Nothing
    Exit Sub

EH:
    p_Error = "TESTS BLOCKED: SetupSandbox failed: " & Err.Description
    m_TestingMode = False
    m_BackendSandboxURL = ""
    Resume CleanExit
End Sub

Private Sub TeardownAll(ByVal p_Db As DAO.Database)
    On Error Resume Next
    p_Db.Execute "DELETE FROM tbDatosPCSUB WHERE idSolicitud >= " & TEST_ID_BASE & " AND idSolicitud <= " & TEST_ID_TOP, dbFailOnError
    p_Db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud >= " & TEST_ID_BASE & " AND idSolicitud <= " & TEST_ID_TOP, dbFailOnError
    p_Db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente >= " & TEST_ID_BASE & " AND IDExpediente <= " & TEST_ID_TOP, dbFailOnError
    On Error GoTo 0
End Sub

' ============================================================================
' SEED HELPERS — schema-first (only Required/NotNull columns from ERD)
'   tbDatosPCSUB.Required: idDatosPCSUB, idSolicitud, refContratoInspeccionOficial
'   tbSolicitudes.Required: idSolicitud, idExpediente, tipoSolicitud,
'     codigoSolicitud, idEstadoInterno, fechaCreacion, usuarioCreacion,
'     revisionCalidadEstado
'   TbExpedientes.Required: IDExpediente
' ============================================================================

Private Sub SeedExpedienteSolicitud(ByVal p_Db As DAO.Database, ByVal p_FixtureId As Long, ByVal p_IdEstadoInterno As Long)
    Dim sql As String
    sql = "INSERT INTO TbExpedientes (IDExpediente) VALUES (" & p_FixtureId & ")"
    p_Db.Execute sql, dbFailOnError

    sql = "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, " & _
          "codigoSolicitud, idEstadoInterno, fechaCreacion, usuarioCreacion, " & _
          "revisionCalidadEstado) VALUES (" & p_FixtureId & ", " & p_FixtureId & ", " & _
          "'PC_SUB', 'PCSUB-STRICT-" & p_FixtureId & "', " & p_IdEstadoInterno & ", Now(), " & _
          "'TestPCSUBStrict', 'PENDIENTE')"
    p_Db.Execute sql, dbFailOnError
End Sub

Private Sub SeedEmptyPCSUB(ByVal p_Db As DAO.Database, ByVal p_FixtureId As Long)
    Dim sql As String
    sql = "INSERT INTO tbDatosPCSUB (idDatosPCSUB, idSolicitud, " & _
          "refContratoInspeccionOficial) VALUES (" & p_FixtureId & ", " & _
          p_FixtureId & ", 'REF-STRICT-" & p_FixtureId & "')"
    p_Db.Execute sql, dbFailOnError
End Sub

' SeedPCSUBRefContratoOnly: minimal seed that satisfies the schema-first rule
' for tests that only need a tbDatosPCSUB row + parent graph (idEstadoInterno
' must be 2 = estadoRegistro so GuardarDatosGenerales does not auto-trigger
' a workflow transition during the field-persistence assertion).
Private Sub SeedMinimal(ByVal p_Db As DAO.Database, ByVal p_FixtureId As Long)
    Call SeedExpedienteSolicitud(p_Db, p_FixtureId, 2)  ' estadoRegistro = 2
    Call SeedEmptyPCSUB(p_Db, p_FixtureId)
End Sub

Private Function CountPCSUBRows(ByVal p_Db As DAO.Database, ByVal p_IdSol As Long) As Long
    Dim rs As DAO.Recordset
    On Error Resume Next
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS n FROM tbDatosPCSUB WHERE idSolicitud = " & p_IdSol)
    If Not rs Is Nothing Then
        CountPCSUBRows = Nz(rs!n, 0)
        rs.Close
    Else
        CountPCSUBRows = 0
    End If
    Set rs = Nothing
End Function

' ============================================================================
' BR-001 — DatosGenerales requires 5 fields; EsDatosGeneralesCompleta check
' ============================================================================

Public Function Test_PCSUB_BR001_GuardarDatosGenerales_PersistsAllFiveRequiredFields() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(7)
    Dim setupError As String
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim rs As DAO.Recordset

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR001_GuardarDatosGenerales_PersistsAllFiveRequiredFields = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR001_ID_HAPPY
    Call SeedMinimal(db, idSol)
    logs(0) = "1. Sandbox OK; seeded minimal fixture idSol=" & idSol

    countBefore = CountPCSUBRows(db, idSol)
    If countBefore <> 1 Then
        logs(1) = "2. FAILED: expected 1 seeded row, got " & countBefore
        Test_PCSUB_BR001_GuardarDatosGenerales_PersistsAllFiveRequiredFields = TestHelper.BuildJsonFail("fixture setup: expected 1 row, got " & countBefore, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. countBefore=" & countBefore & " (expected 1)"

    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.refContratoInspeccionOficial = "BR001-REF-NEW"
    pcsub.refSubSuministrador = "BR001-SUB-NEW"
    pcsub.SubsuministradorNombreDir = "BR001 Sub NomDir"
    pcsub.denominacionContrato = "BR001 Denominacion"
    pcsub.objetoContrato = "BR001 Objeto"

    svc.GuardarDatosGenerales pcsub, db
    logs(2) = "3. GuardarDatosGenerales called with injected db"

    countAfter = CountPCSUBRows(db, idSol)
    If countAfter <> 1 Then
        logs(3) = "4. FAILED: row count changed after UPDATE: before=" & countBefore & " after=" & countAfter
        Test_PCSUB_BR001_GuardarDatosGenerales_PersistsAllFiveRequiredFields = TestHelper.BuildJsonFail("UPDATE changed row count", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=" & countAfter & " (expected 1)"

    Set rs = db.OpenRecordset( _
        "SELECT refContratoInspeccionOficial, refSubSuministrador, SubsuministradorNombreDir, " & _
        "denominacionContrato, objetoContrato FROM tbDatosPCSUB WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If rs.EOF Then
        logs(4) = "5. FAILED: row missing after update"
        Test_PCSUB_BR001_GuardarDatosGenerales_PersistsAllFiveRequiredFields = TestHelper.BuildJsonFail("row missing", logs)
        GoTo CleanExit
    End If

    If Nz(rs!refContratoInspeccionOficial, "") <> "BR001-REF-NEW" _
       Or Nz(rs!refSubSuministrador, "") <> "BR001-SUB-NEW" _
       Or Nz(rs!SubsuministradorNombreDir, "") <> "BR001 Sub NomDir" _
       Or Nz(rs!denominacionContrato, "") <> "BR001 Denominacion" _
       Or Nz(rs!objetoContrato, "") <> "BR001 Objeto" Then
        logs(4) = "5. FAILED: at least one of 5 required fields not persisted"
        Test_PCSUB_BR001_GuardarDatosGenerales_PersistsAllFiveRequiredFields = TestHelper.BuildJsonFail("one or more required fields not persisted", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. All 5 BR-001 required fields persisted through injected db"

    logs(5) = "6. complete"
    logs(6) = "7. PASS"
    Test_PCSUB_BR001_GuardarDatosGenerales_PersistsAllFiveRequiredFields = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR001_GuardarDatosGenerales_PersistsAllFiveRequiredFields: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR001_GuardarDatosGenerales_PersistsAllFiveRequiredFields = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR001_GuardarDatosGenerales_RejectsEmptyRefContrato() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim setupError As String
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim errNumber As Long

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR001_GuardarDatosGenerales_RejectsEmptyRefContrato = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR001_ID_REJ_REF
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.refContratoInspeccionOficial = ""  ' violates BR-001 required field
    pcsub.refSubSuministrador = "BR001-SUB"
    pcsub.SubsuministradorNombreDir = "BR001 Sub"
    pcsub.denominacionContrato = "BR001 Den"
    pcsub.objetoContrato = "BR001 Obj"
    logs(0) = "1. Seeded idSol=" & idSol & " with empty refContratoInspeccionOficial"

    countBefore = CountPCSUBRows(db, idSol)
    If countBefore <> 1 Then
        logs(1) = "2. FAILED: expected 1 seeded row, got " & countBefore
        Test_PCSUB_BR001_GuardarDatosGenerales_RejectsEmptyRefContrato = TestHelper.BuildJsonFail("fixture setup", logs)
        GoTo CleanExit
    End If

    On Error Resume Next
    svc.GuardarDatosGenerales pcsub, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 513 Then
        logs(2) = "2. FAILED: expected Err 513, got " & errNumber
        Test_PCSUB_BR001_GuardarDatosGenerales_RejectsEmptyRefContrato = TestHelper.BuildJsonFail("expected Err 513, got " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(2) = "2. GuardarDatosGenerales raised Err 513 as expected (empty refContrato)"

    countAfter = CountPCSUBRows(db, idSol)
    If countAfter <> countBefore Then
        logs(3) = "3. FAILED: row count changed on validation error: before=" & countBefore & " after=" & countAfter
        Test_PCSUB_BR001_GuardarDatosGenerales_RejectsEmptyRefContrato = TestHelper.BuildJsonFail("validation error changed cardinality", logs)
        GoTo CleanExit
    End If
    logs(3) = "3. countAfter=" & countAfter & " (expected " & countBefore & " — no side effect)"

    logs(4) = "4. complete"
    logs(5) = "5. PASS"
    Test_PCSUB_BR001_GuardarDatosGenerales_RejectsEmptyRefContrato = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR001_GuardarDatosGenerales_RejectsEmptyRefContrato: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR001_GuardarDatosGenerales_RejectsEmptyRefContrato = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR001_EsDatosGeneralesCompleta_TrueForCompleteRow() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim result As Boolean

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR001_EsDatosGeneralesCompleta_TrueForCompleteRow = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR001_ID_COMP
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.refContratoInspeccionOficial = "BR001-COMP-REF"
    pcsub.refSubSuministrador = "BR001-COMP-SUB"
    pcsub.SubsuministradorNombreDir = "BR001-COMP-SUB-DIR"
    pcsub.denominacionContrato = "BR001-COMP-DEN"
    pcsub.objetoContrato = "BR001-COMP-OBJ"
    svc.GuardarDatosGenerales pcsub, db
    logs(0) = "1. Seeded complete general fields on idSol=" & idSol

    result = svc.EsDatosGeneralesCompleta(idSol, db)
    If Not result Then
        logs(1) = "2. FAILED: EsDatosGeneralesCompleta returned False on complete row"
        Test_PCSUB_BR001_EsDatosGeneralesCompleta_TrueForCompleteRow = TestHelper.BuildJsonFail("EsDatosGeneralesCompleta=False on complete row", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. EsDatosGeneralesCompleta=True as expected for complete row"

    logs(2) = "3. complete"
    logs(3) = "4. PASS"
    Test_PCSUB_BR001_EsDatosGeneralesCompleta_TrueForCompleteRow = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR001_EsDatosGeneralesCompleta_TrueForCompleteRow: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR001_EsDatosGeneralesCompleta_TrueForCompleteRow = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR001_EsDatosGeneralesCompleta_FalseWhenOneFieldMissing() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim result As Boolean

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR001_EsDatosGeneralesCompleta_FalseWhenOneFieldMissing = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR001_ID_INCOMP
    ' Seed only the parent graph; create a DatosPCSUB via the repository
    ' (which bypasses GuardarDatosGenerales validation) to deliberately
    ' leave objetoContrato empty. We cannot use GuardarDatosGenerales here
    ' because the service rejects empty required fields by design.
    Call SeedExpedienteSolicitud(db, idSol, 2)
    Set pcsub = New DatosPCSUB
    pcsub.idSolicitud = idSol
    pcsub.refContratoInspeccionOficial = "BR001-INCOMP-REF"
    pcsub.refSubSuministrador = "BR001-INCOMP-SUB"
    pcsub.SubsuministradorNombreDir = "BR001-INCOMP-SUB-DIR"
    pcsub.denominacionContrato = "BR001-INCOMP-DEN"
    pcsub.objetoContrato = ""  ' violates BR-001
    DatosPCSUBRepositorio.Guardar pcsub, db
    logs(0) = "1. Seeded idSol=" & idSol & " with objetoContrato empty (BR-001 violation, via repo to bypass service validation)"

    result = svc.EsDatosGeneralesCompleta(idSol, db)
    If result Then
        logs(1) = "2. FAILED: EsDatosGeneralesCompleta returned True when objetoContrato is empty"
        Test_PCSUB_BR001_EsDatosGeneralesCompleta_FalseWhenOneFieldMissing = TestHelper.BuildJsonFail("EsDatosGeneralesCompleta=True with empty objetoContrato", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. EsDatosGeneralesCompleta=False as expected (objetoContrato empty)"

    logs(2) = "3. complete"
    logs(3) = "4. PASS"
    Test_PCSUB_BR001_EsDatosGeneralesCompleta_FalseWhenOneFieldMissing = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR001_EsDatosGeneralesCompleta_FalseWhenOneFieldMissing: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR001_EsDatosGeneralesCompleta_FalseWhenOneFieldMissing = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ============================================================================
' BR-002 — Propuesta requires material afectado + descripcion propuesta
' ============================================================================

Public Function Test_PCSUB_BR002_GuardarPropuesta_PersistsMaterialYPropuesta() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim setupError As String
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim rs As DAO.Recordset

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR002_GuardarPropuesta_PersistsMaterialYPropuesta = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR002_ID_HAPPY
    Call SeedMinimal(db, idSol)
    countBefore = CountPCSUBRows(db, idSol)
    If countBefore <> 1 Then
        logs(0) = "1. FAILED: expected 1 row, got " & countBefore
        Test_PCSUB_BR002_GuardarPropuesta_PersistsMaterialYPropuesta = TestHelper.BuildJsonFail("fixture setup", logs)
        GoTo CleanExit
    End If
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.descripcionMaterialAfectado = "BR002-MAT"
    pcsub.descripcionPropuestaCambio = "BR002-PROP"
    pcsub.numPlanoEspecificacion = "BR002-PLANO"
    pcsub.descripcionPropuestaCambioCont = "BR002-CONT"

    svc.GuardarPropuesta pcsub, db
    logs(1) = "2. GuardarPropuesta called with injected db"

    countAfter = CountPCSUBRows(db, idSol)
    If countAfter <> 1 Then
        logs(2) = "3. FAILED: row count changed: before=" & countBefore & " after=" & countAfter
        Test_PCSUB_BR002_GuardarPropuesta_PersistsMaterialYPropuesta = TestHelper.BuildJsonFail("row count changed", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected 1)"

    Set rs = db.OpenRecordset( _
        "SELECT descripcionMaterialAfectado, descripcionPropuestaCambio, numPlanoEspecificacion, " & _
        "descripcionPropuestaCambioCont FROM tbDatosPCSUB WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If Nz(rs!descripcionMaterialAfectado, "") <> "BR002-MAT" _
       Or Nz(rs!descripcionPropuestaCambio, "") <> "BR002-PROP" _
       Or Nz(rs!numPlanoEspecificacion, "") <> "BR002-PLANO" _
       Or Nz(rs!descripcionPropuestaCambioCont, "") <> "BR002-CONT" Then
        logs(3) = "4. FAILED: propuesta fields not persisted"
        Test_PCSUB_BR002_GuardarPropuesta_PersistsMaterialYPropuesta = TestHelper.BuildJsonFail("propuesta fields not persisted", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. All 4 BR-002 propuesta fields persisted"

    logs(4) = "5. complete"
    logs(5) = "6. PASS"
    Test_PCSUB_BR002_GuardarPropuesta_PersistsMaterialYPropuesta = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR002_GuardarPropuesta_PersistsMaterialYPropuesta: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR002_GuardarPropuesta_PersistsMaterialYPropuesta = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR002_GuardarPropuesta_RejectsEmptyMaterialAfectado() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim errNumber As Long

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR002_GuardarPropuesta_RejectsEmptyMaterialAfectado = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR002_ID_NO_MATERIAL
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.descripcionMaterialAfectado = ""  ' violates BR-002
    pcsub.descripcionPropuestaCambio = "BR002-PROP"
    countBefore = CountPCSUBRows(db, idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " with empty material, countBefore=" & countBefore

    On Error Resume Next
    svc.GuardarPropuesta pcsub, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 513 Then
        logs(1) = "2. FAILED: expected Err 513, got " & errNumber
        Test_PCSUB_BR002_GuardarPropuesta_RejectsEmptyMaterialAfectado = TestHelper.BuildJsonFail("expected Err 513, got " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarPropuesta raised Err 513 as expected (empty material)"

    countAfter = CountPCSUBRows(db, idSol)
    If countAfter <> countBefore Then
        logs(2) = "3. FAILED: cardinality changed on validation error"
        Test_PCSUB_BR002_GuardarPropuesta_RejectsEmptyMaterialAfectado = TestHelper.BuildJsonFail("cardinality changed", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (no side effect)"

    logs(3) = "4. complete"
    logs(4) = "5. PASS"
    Test_PCSUB_BR002_GuardarPropuesta_RejectsEmptyMaterialAfectado = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR002_GuardarPropuesta_RejectsEmptyMaterialAfectado: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR002_GuardarPropuesta_RejectsEmptyMaterialAfectado = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR002_GuardarPropuesta_RejectsEmptyDescripcionPropuesta() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim errNumber As Long

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR002_GuardarPropuesta_RejectsEmptyDescripcionPropuesta = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR002_ID_NO_PROPUESTA
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.descripcionMaterialAfectado = "BR002-MAT"
    pcsub.descripcionPropuestaCambio = ""  ' violates BR-002
    logs(0) = "1. Seeded idSol=" & idSol & " with empty propuesta"

    On Error Resume Next
    svc.GuardarPropuesta pcsub, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 513 Then
        logs(1) = "2. FAILED: expected Err 513, got " & errNumber
        Test_PCSUB_BR002_GuardarPropuesta_RejectsEmptyDescripcionPropuesta = TestHelper.BuildJsonFail("expected Err 513, got " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarPropuesta raised Err 513 as expected (empty propuesta)"

    logs(2) = "3. complete"
    logs(3) = "4. PASS"
    Test_PCSUB_BR002_GuardarPropuesta_RejectsEmptyDescripcionPropuesta = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR002_GuardarPropuesta_RejectsEmptyDescripcionPropuesta: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR002_GuardarPropuesta_RejectsEmptyDescripcionPropuesta = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR002_EsDetalleCompleto_TrueForCompleteDetail() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim result As Boolean

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR002_EsDetalleCompleto_TrueForCompleteDetail = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR002_ID_COMP
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.descripcionMaterialAfectado = "BR002-DET-MAT"
    pcsub.descripcionPropuestaCambio = "BR002-DET-PROP"
    svc.GuardarPropuesta pcsub, db
    logs(0) = "1. Seeded complete detail on idSol=" & idSol

    result = svc.EsDetalleCompleto(idSol, db)
    If Not result Then
        logs(1) = "2. FAILED: EsDetalleCompleto=False on complete detail"
        Test_PCSUB_BR002_EsDetalleCompleto_TrueForCompleteDetail = TestHelper.BuildJsonFail("EsDetalleCompleto=False", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. EsDetalleCompleto=True as expected"

    logs(2) = "3. complete"
    logs(3) = "4. PASS"
    Test_PCSUB_BR002_EsDetalleCompleto_TrueForCompleteDetail = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR002_EsDetalleCompleto_TrueForCompleteDetail: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR002_EsDetalleCompleto_TrueForCompleteDetail = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ============================================================================
' BR-003 — Impacto: at least 1 motivo; motivoOtros requires detalle;
'           detalle without motivoOtros rejected
' ============================================================================

Public Function Test_PCSUB_BR003_GuardarImpacto_RequiresAtLeastOneMotivo() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim errNumber As Long

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR003_GuardarImpacto_RequiresAtLeastOneMotivo = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR003_ID_NO_MOTIVO
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    ' All motive flags left False (no motivo marked)
    pcsub.incidenciaCoste = "AUMENTARÁ"
    pcsub.incidenciaPlazo = "NO VARIARÁ"
    pcsub.impactoClasificacion = "MAYOR"
    pcsub.CambioAfectaAMaterial = "Material por entregar"
    logs(0) = "1. Seeded idSol=" & idSol & " with all motive flags False"

    On Error Resume Next
    svc.GuardarImpacto pcsub, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 513 Then
        logs(1) = "2. FAILED: expected Err 513, got " & errNumber
        Test_PCSUB_BR003_GuardarImpacto_RequiresAtLeastOneMotivo = TestHelper.BuildJsonFail("expected Err 513, got " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarImpacto raised Err 513 as expected (no motivo)"

    logs(2) = "3. complete"
    logs(3) = "4. PASS"
    Test_PCSUB_BR003_GuardarImpacto_RequiresAtLeastOneMotivo = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR003_GuardarImpacto_RequiresAtLeastOneMotivo: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR003_GuardarImpacto_RequiresAtLeastOneMotivo = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR003_GuardarImpacto_OtrosDetalle_RequiresOthersMarked() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim errNumber As Long

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR003_GuardarImpacto_OtrosDetalle_RequiresOthersMarked = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR003_ID_OTROS_DETALLE_NO_FLAG
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.motivoMejorarCapacidad = True   ' at least 1 motivo so motivos check passes
    pcsub.motivoOtros = False             ' but motivoOtros NOT marked
    pcsub.motivoOtrosDetalle = "Detalle sin marcar Otros"  ' violation
    pcsub.incidenciaCoste = "AUMENTARÁ"
    pcsub.incidenciaPlazo = "NO VARIARÁ"
    pcsub.impactoClasificacion = "MAYOR"
    pcsub.CambioAfectaAMaterial = "Material por entregar"
    logs(0) = "1. Seeded idSol=" & idSol & " motivoOtros=False but detalle populated"

    On Error Resume Next
    svc.GuardarImpacto pcsub, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 513 Then
        logs(1) = "2. FAILED: expected Err 513, got " & errNumber
        Test_PCSUB_BR003_GuardarImpacto_OtrosDetalle_RequiresOthersMarked = TestHelper.BuildJsonFail("expected Err 513, got " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarImpacto raised Err 513 as expected (detalle without Otros flag)"

    logs(2) = "3. complete"
    logs(3) = "4. PASS"
    Test_PCSUB_BR003_GuardarImpacto_OtrosDetalle_RequiresOthersMarked = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR003_GuardarImpacto_OtrosDetalle_RequiresOthersMarked: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR003_GuardarImpacto_OtrosDetalle_RequiresOthersMarked = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR003_EsMotivosCompleto_TrueForSeededMotivo() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim result As Boolean

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR003_EsMotivosCompleto_TrueForSeededMotivo = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR003_ID_COMP
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.motivoMejorarSeguridad = True
    pcsub.incidenciaCoste = "AUMENTARÁ"
    pcsub.incidenciaPlazo = "NO VARIARÁ"
    pcsub.impactoClasificacion = "MAYOR"
    pcsub.CambioAfectaAMaterial = "Material por entregar"
    svc.GuardarImpacto pcsub, db
    logs(0) = "1. Seeded motivoMejorarSeguridad=True on idSol=" & idSol

    result = svc.EsMotivosCompleto(idSol, db)
    If Not result Then
        logs(1) = "2. FAILED: EsMotivosCompleto=False with motivoMejorarSeguridad=True"
        Test_PCSUB_BR003_EsMotivosCompleto_TrueForSeededMotivo = TestHelper.BuildJsonFail("EsMotivosCompleto=False", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. EsMotivosCompleto=True as expected"

    logs(2) = "3. complete"
    logs(3) = "4. PASS"
    Test_PCSUB_BR003_EsMotivosCompleto_TrueForSeededMotivo = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR003_EsMotivosCompleto_TrueForSeededMotivo: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR003_EsMotivosCompleto_TrueForSeededMotivo = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ============================================================================
' BR-004 — IncidenciaCoste/Plazo enum (AUMENTARÁ|DISMINUIRÁ|NO VARIARÁ)
' ============================================================================

Public Function Test_PCSUB_BR004_GuardarImpacto_RejectsIncidenciaCosteInvalida() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim errNumber As Long

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR004_GuardarImpacto_RejectsIncidenciaCosteInvalida = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR004_ID_BAD_COSTE
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.motivoMejorarSeguridad = True
    pcsub.incidenciaCoste = "VALOR INVÁLIDO"  ' violates BR-004 enum
    pcsub.incidenciaPlazo = "NO VARIARÁ"
    pcsub.impactoClasificacion = "MAYOR"
    pcsub.CambioAfectaAMaterial = "Material por entregar"
    logs(0) = "1. Seeded idSol=" & idSol & " with invalid incidenciaCoste"

    On Error Resume Next
    svc.GuardarImpacto pcsub, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 513 Then
        logs(1) = "2. FAILED: expected Err 513, got " & errNumber
        Test_PCSUB_BR004_GuardarImpacto_RejectsIncidenciaCosteInvalida = TestHelper.BuildJsonFail("expected Err 513, got " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarImpacto raised Err 513 as expected (invalida incidenciaCoste)"

    logs(2) = "3. complete"
    logs(3) = "4. PASS"
    Test_PCSUB_BR004_GuardarImpacto_RejectsIncidenciaCosteInvalida = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR004_GuardarImpacto_RejectsIncidenciaCosteInvalida: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR004_GuardarImpacto_RejectsIncidenciaCosteInvalida = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR004_GuardarImpacto_RejectsIncidenciaPlazoInvalida() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim errNumber As Long

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR004_GuardarImpacto_RejectsIncidenciaPlazoInvalida = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR004_ID_BAD_PLAZO
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.motivoMejorarSeguridad = True
    pcsub.incidenciaCoste = "AUMENTARÁ"
    pcsub.incidenciaPlazo = "VALOR INVÁLIDO"  ' violates BR-004 enum
    pcsub.impactoClasificacion = "MAYOR"
    pcsub.CambioAfectaAMaterial = "Material por entregar"
    logs(0) = "1. Seeded idSol=" & idSol & " with invalid incidenciaPlazo"

    On Error Resume Next
    svc.GuardarImpacto pcsub, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 513 Then
        logs(1) = "2. FAILED: expected Err 513, got " & errNumber
        Test_PCSUB_BR004_GuardarImpacto_RejectsIncidenciaPlazoInvalida = TestHelper.BuildJsonFail("expected Err 513, got " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarImpacto raised Err 513 as expected (invalida incidenciaPlazo)"

    logs(2) = "3. complete"
    logs(3) = "4. PASS"
    Test_PCSUB_BR004_GuardarImpacto_RejectsIncidenciaPlazoInvalida = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR004_GuardarImpacto_RejectsIncidenciaPlazoInvalida: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR004_GuardarImpacto_RejectsIncidenciaPlazoInvalida = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ============================================================================
' BR-005 — ImpactoClasificacion enum (MAYOR|MENOR)
' ============================================================================

Public Function Test_PCSUB_BR005_GuardarImpacto_RejectsClasificacionInvalida() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim errNumber As Long

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR005_GuardarImpacto_RejectsClasificacionInvalida = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR005_ID_BAD_CLASIF
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.motivoMejorarSeguridad = True
    pcsub.incidenciaCoste = "AUMENTARÁ"
    pcsub.incidenciaPlazo = "NO VARIARÁ"
    pcsub.impactoClasificacion = "INVALIDA"  ' violates BR-005 enum
    pcsub.CambioAfectaAMaterial = "Material por entregar"
    logs(0) = "1. Seeded idSol=" & idSol & " with invalid impactoClasificacion"

    On Error Resume Next
    svc.GuardarImpacto pcsub, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 513 Then
        logs(1) = "2. FAILED: expected Err 513, got " & errNumber
        Test_PCSUB_BR005_GuardarImpacto_RejectsClasificacionInvalida = TestHelper.BuildJsonFail("expected Err 513, got " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarImpacto raised Err 513 as expected (invalida impactoClasificacion)"

    logs(2) = "3. complete"
    logs(3) = "4. PASS"
    Test_PCSUB_BR005_GuardarImpacto_RejectsClasificacionInvalida = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR005_GuardarImpacto_RejectsClasificacionInvalida: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR005_GuardarImpacto_RejectsClasificacionInvalida = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ============================================================================
' BR-006 — CambioAfectaAMaterial enum (Material ya entregado|por entregar)
' ============================================================================

Public Function Test_PCSUB_BR006_GuardarImpacto_RejectsCambioAfectaAMaterialInvalido() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim errNumber As Long

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR006_GuardarImpacto_RejectsCambioAfectaAMaterialInvalido = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR006_ID_BAD_MATERIAL
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.motivoMejorarSeguridad = True
    pcsub.incidenciaCoste = "AUMENTARÁ"
    pcsub.incidenciaPlazo = "NO VARIARÁ"
    pcsub.impactoClasificacion = "MAYOR"
    pcsub.CambioAfectaAMaterial = "VALOR INVÁLIDO"  ' violates BR-006 enum
    logs(0) = "1. Seeded idSol=" & idSol & " with invalid CambioAfectaAMaterial"

    On Error Resume Next
    svc.GuardarImpacto pcsub, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 513 Then
        logs(1) = "2. FAILED: expected Err 513, got " & errNumber
        Test_PCSUB_BR006_GuardarImpacto_RejectsCambioAfectaAMaterialInvalido = TestHelper.BuildJsonFail("expected Err 513, got " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarImpacto raised Err 513 as expected (invalid CambioAfectaAMaterial)"

    logs(2) = "3. complete"
    logs(3) = "4. PASS"
    Test_PCSUB_BR006_GuardarImpacto_RejectsCambioAfectaAMaterialInvalido = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR006_GuardarImpacto_RejectsCambioAfectaAMaterialInvalido: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR006_GuardarImpacto_RejectsCambioAfectaAMaterialInvalido = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ============================================================================
' BR-007 — AprobacionSuministrador requires two signer names
' ============================================================================

Public Function Test_PCSUB_BR007_GuardarAprobacionSuministrador_PersistsBothSigners() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim setupError As String
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim rs As DAO.Recordset

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR007_GuardarAprobacionSuministrador_PersistsBothSigners = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR007_ID_HAPPY
    Call SeedMinimal(db, idSol)
    countBefore = CountPCSUBRows(db, idSol)
    If countBefore <> 1 Then
        logs(0) = "1. FAILED: expected 1 row, got " & countBefore
        Test_PCSUB_BR007_GuardarAprobacionSuministrador_PersistsBothSigners = TestHelper.BuildJsonFail("fixture setup", logs)
        GoTo CleanExit
    End If
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.firmaOficinaTecnicaSubSuministradorNombre = "BR007-OTICINA"
    pcsub.firmaRepSubSuministradorNombre = "BR007-REPSUB"

    svc.GuardarAprobacionSuministrador pcsub, db
    logs(1) = "2. GuardarAprobacionSuministrador called with injected db"

    countAfter = CountPCSUBRows(db, idSol)
    If countAfter <> 1 Then
        logs(2) = "3. FAILED: row count changed: before=" & countBefore & " after=" & countAfter
        Test_PCSUB_BR007_GuardarAprobacionSuministrador_PersistsBothSigners = TestHelper.BuildJsonFail("row count changed", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected 1)"

    Set rs = db.OpenRecordset( _
        "SELECT firmaOficinaTecnicaSubSuministradorNombre, firmaRepSubSuministradorNombre " & _
        "FROM tbDatosPCSUB WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If Nz(rs!firmaOficinaTecnicaSubSuministradorNombre, "") <> "BR007-OTICINA" _
       Or Nz(rs!firmaRepSubSuministradorNombre, "") <> "BR007-REPSUB" Then
        logs(3) = "4. FAILED: signer names not persisted through injected db"
        Test_PCSUB_BR007_GuardarAprobacionSuministrador_PersistsBothSigners = TestHelper.BuildJsonFail("signer names not persisted", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. Both signer names persisted through injected db"

    logs(4) = "5. complete"
    logs(5) = "6. PASS"
    Test_PCSUB_BR007_GuardarAprobacionSuministrador_PersistsBothSigners = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR007_GuardarAprobacionSuministrador_PersistsBothSigners: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR007_GuardarAprobacionSuministrador_PersistsBothSigners = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR007_GuardarAprobacionSuministrador_RejectsEmptySigner1() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim errNumber As Long

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR007_GuardarAprobacionSuministrador_RejectsEmptySigner1 = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR007_ID_REJ_S1
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.firmaOficinaTecnicaSubSuministradorNombre = ""  ' violates BR-007
    pcsub.firmaRepSubSuministradorNombre = "BR007-S2"
    logs(0) = "1. Seeded idSol=" & idSol & " with empty signer 1"

    On Error Resume Next
    svc.GuardarAprobacionSuministrador pcsub, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 513 Then
        logs(1) = "2. FAILED: expected Err 513, got " & errNumber
        Test_PCSUB_BR007_GuardarAprobacionSuministrador_RejectsEmptySigner1 = TestHelper.BuildJsonFail("expected Err 513, got " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarAprobacionSuministrador raised Err 513 as expected (empty signer 1)"

    logs(2) = "3. complete"
    logs(3) = "4. PASS"
    Test_PCSUB_BR007_GuardarAprobacionSuministrador_RejectsEmptySigner1 = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR007_GuardarAprobacionSuministrador_RejectsEmptySigner1: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR007_GuardarAprobacionSuministrador_RejectsEmptySigner1 = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR007_GuardarAprobacionSuministrador_RejectsEmptySigner2() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim errNumber As Long

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR007_GuardarAprobacionSuministrador_RejectsEmptySigner2 = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR007_ID_REJ_S2
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.firmaOficinaTecnicaSubSuministradorNombre = "BR007-S1"
    pcsub.firmaRepSubSuministradorNombre = ""  ' violates BR-007
    logs(0) = "1. Seeded idSol=" & idSol & " with empty signer 2"

    On Error Resume Next
    svc.GuardarAprobacionSuministrador pcsub, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 513 Then
        logs(1) = "2. FAILED: expected Err 513, got " & errNumber
        Test_PCSUB_BR007_GuardarAprobacionSuministrador_RejectsEmptySigner2 = TestHelper.BuildJsonFail("expected Err 513, got " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarAprobacionSuministrador raised Err 513 as expected (empty signer 2)"

    logs(2) = "3. complete"
    logs(3) = "4. PASS"
    Test_PCSUB_BR007_GuardarAprobacionSuministrador_RejectsEmptySigner2 = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR007_GuardarAprobacionSuministrador_RejectsEmptySigner2: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR007_GuardarAprobacionSuministrador_RejectsEmptySigner2 = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ============================================================================
' BR-008 — DictamenRAC: non-RECHAZADO requires racCodigo; RECHAZADO allows empty
' ============================================================================

Public Function Test_PCSUB_BR008_GuardarDictamenRAC_PersistsValidRAC() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim setupError As String
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim rs As DAO.Recordset

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR008_GuardarDictamenRAC_PersistsValidRAC = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR008_ID_HAPPY
    Call SeedMinimal(db, idSol)
    countBefore = CountPCSUBRows(db, idSol)
    If countBefore <> 1 Then
        logs(0) = "1. FAILED: expected 1 row, got " & countBefore
        Test_PCSUB_BR008_GuardarDictamenRAC_PersistsValidRAC = TestHelper.BuildJsonFail("fixture setup", logs)
        GoTo CleanExit
    End If
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.racCodigo = "BR008-RAC-CODE"
    pcsub.racDecision = "APROBADO"
    pcsub.racNombre = "BR008-RAC-NAME"
    pcsub.observacionesRAC = "BR008-RAC-OBS"

    svc.GuardarDictamenRAC pcsub, db
    logs(1) = "2. GuardarDictamenRAC called with injected db (APROBADO+code)"

    countAfter = CountPCSUBRows(db, idSol)
    If countAfter <> 1 Then
        logs(2) = "3. FAILED: row count changed: before=" & countBefore & " after=" & countAfter
        Test_PCSUB_BR008_GuardarDictamenRAC_PersistsValidRAC = TestHelper.BuildJsonFail("row count changed", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected 1)"

    Set rs = db.OpenRecordset( _
        "SELECT racCodigo, racDecision, racNombre, observacionesRAC FROM tbDatosPCSUB WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If Nz(rs!racCodigo, "") <> "BR008-RAC-CODE" _
       Or Nz(rs!racDecision, "") <> "APROBADO" _
       Or Nz(rs!racNombre, "") <> "BR008-RAC-NAME" _
       Or Nz(rs!observacionesRAC, "") <> "BR008-RAC-OBS" Then
        logs(3) = "4. FAILED: RAC fields not persisted"
        Test_PCSUB_BR008_GuardarDictamenRAC_PersistsValidRAC = TestHelper.BuildJsonFail("RAC fields not persisted", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. All 4 RAC fields persisted through injected db"

    logs(4) = "5. complete"
    logs(5) = "6. PASS"
    Test_PCSUB_BR008_GuardarDictamenRAC_PersistsValidRAC = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR008_GuardarDictamenRAC_PersistsValidRAC: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR008_GuardarDictamenRAC_PersistsValidRAC = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR008_GuardarDictamenRAC_RejectedDecision_AcceptsEmptyCodigo() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim errNumber As Long

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR008_GuardarDictamenRAC_RejectedDecision_AcceptsEmptyCodigo = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR008_ID_REJ_EMPTY
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.racCodigo = ""  ' empty OK for RECHAZADO (BR-008 exception)
    pcsub.racDecision = "RECHAZADO"
    pcsub.racNombre = "BR008-REJ"
    logs(0) = "1. Seeded idSol=" & idSol & " racDecision=RECHAZADO racCodigo=empty"

    On Error Resume Next
    svc.GuardarDictamenRAC pcsub, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: expected NO error for RECHAZADO+empty code, got Err " & errNumber
        Test_PCSUB_BR008_GuardarDictamenRAC_RejectedDecision_AcceptsEmptyCodigo = TestHelper.BuildJsonFail("unexpected Err " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarDictamenRAC succeeded with RECHAZADO+empty code (BR-008 exception honored)"

    logs(2) = "3. complete"
    logs(3) = "4. PASS"
    Test_PCSUB_BR008_GuardarDictamenRAC_RejectedDecision_AcceptsEmptyCodigo = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR008_GuardarDictamenRAC_RejectedDecision_AcceptsEmptyCodigo: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR008_GuardarDictamenRAC_RejectedDecision_AcceptsEmptyCodigo = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR008_GuardarDictamenRAC_NonRejected_RequiresCodigo() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim errNumber As Long

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR008_GuardarDictamenRAC_NonRejected_RequiresCodigo = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR008_ID_REJ_NONREJ
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.racCodigo = ""  ' violates BR-008 (decision is not RECHAZADO)
    pcsub.racDecision = "APROBADO"
    pcsub.racNombre = "BR008-APR"
    logs(0) = "1. Seeded idSol=" & idSol & " racDecision=APROBADO racCodigo=empty (BR-008 violation)"

    On Error Resume Next
    svc.GuardarDictamenRAC pcsub, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 513 Then
        logs(1) = "2. FAILED: expected Err 513, got " & errNumber
        Test_PCSUB_BR008_GuardarDictamenRAC_NonRejected_RequiresCodigo = TestHelper.BuildJsonFail("expected Err 513, got " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarDictamenRAC raised Err 513 as expected (non-RECHAZADO without code)"

    logs(2) = "3. complete"
    logs(3) = "4. PASS"
    Test_PCSUB_BR008_GuardarDictamenRAC_NonRejected_RequiresCodigo = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR008_GuardarDictamenRAC_NonRejected_RequiresCodigo: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR008_GuardarDictamenRAC_NonRejected_RequiresCodigo = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ============================================================================
' BR-009 — DecisionFinal requires decisionFinal
' ============================================================================

Public Function Test_PCSUB_BR009_GuardarDecisionFinal_PersistsValidDecision() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim setupError As String
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim rs As DAO.Recordset

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR009_GuardarDecisionFinal_PersistsValidDecision = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR009_ID_HAPPY
    Call SeedMinimal(db, idSol)
    countBefore = CountPCSUBRows(db, idSol)
    If countBefore <> 1 Then
        logs(0) = "1. FAILED: expected 1 row, got " & countBefore
        Test_PCSUB_BR009_GuardarDecisionFinal_PersistsValidDecision = TestHelper.BuildJsonFail("fixture setup", logs)
        GoTo CleanExit
    End If
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.decisionFinal = "BR009-APROBADO"
    pcsub.NombreFirmanteFinal = "BR009-FIRMANTE"
    pcsub.obsDecisionFinal = "BR009-OBS"

    svc.GuardarDecisionFinal pcsub, db
    logs(1) = "2. GuardarDecisionFinal called with injected db"

    countAfter = CountPCSUBRows(db, idSol)
    If countAfter <> 1 Then
        logs(2) = "3. FAILED: row count changed: before=" & countBefore & " after=" & countAfter
        Test_PCSUB_BR009_GuardarDecisionFinal_PersistsValidDecision = TestHelper.BuildJsonFail("row count changed", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected 1)"

    Set rs = db.OpenRecordset( _
        "SELECT decisionFinal, NombreFirmanteFinal, obsDecisionFinal FROM tbDatosPCSUB WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If Nz(rs!decisionFinal, "") <> "BR009-APROBADO" _
       Or Nz(rs!NombreFirmanteFinal, "") <> "BR009-FIRMANTE" _
       Or Nz(rs!obsDecisionFinal, "") <> "BR009-OBS" Then
        logs(3) = "4. FAILED: decision fields not persisted"
        Test_PCSUB_BR009_GuardarDecisionFinal_PersistsValidDecision = TestHelper.BuildJsonFail("decision fields not persisted", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. All 3 decision-final fields persisted"

    logs(4) = "5. complete"
    logs(5) = "6. PASS"
    Test_PCSUB_BR009_GuardarDecisionFinal_PersistsValidDecision = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR009_GuardarDecisionFinal_PersistsValidDecision: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR009_GuardarDecisionFinal_PersistsValidDecision = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR009_GuardarDecisionFinal_RejectsEmptyDecision() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim errNumber As Long

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR009_GuardarDecisionFinal_RejectsEmptyDecision = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR009_ID_EMPTY
    Call SeedMinimal(db, idSol)
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.decisionFinal = ""  ' violates BR-009
    logs(0) = "1. Seeded idSol=" & idSol & " with empty decisionFinal"

    On Error Resume Next
    svc.GuardarDecisionFinal pcsub, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 513 Then
        logs(1) = "2. FAILED: expected Err 513, got " & errNumber
        Test_PCSUB_BR009_GuardarDecisionFinal_RejectsEmptyDecision = TestHelper.BuildJsonFail("expected Err 513, got " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarDecisionFinal raised Err 513 as expected (empty decisionFinal)"

    logs(2) = "3. complete"
    logs(3) = "4. PASS"
    Test_PCSUB_BR009_GuardarDecisionFinal_RejectsEmptyDecision = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR009_GuardarDecisionFinal_RejectsEmptyDecision: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR009_GuardarDecisionFinal_RejectsEmptyDecision = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ============================================================================
' BR-010 — Limpiar* side-effects: only the target block is cleared
' ============================================================================

Public Function Test_PCSUB_BR010_PurgaTecnica_ClearsOnlyTecnicaFields() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim rs As DAO.Recordset

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR010_PurgaTecnica_ClearsOnlyTecnicaFields = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR010_ID_PURGA
    Call SeedMinimal(db, idSol)
    ' Seed full tecnica fields on the entity and persist via GuardarImpacto
    Dim pcsub As DatosPCSUB
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.descripcionMaterialAfectado = "BR010-MAT"
    pcsub.descripcionPropuestaCambio = "BR010-PROP"
    pcsub.numPlanoEspecificacion = "BR010-PLANO"
    pcsub.motivoMejorarSeguridad = True
    pcsub.incidenciaCoste = "AUMENTARÁ"
    pcsub.incidenciaPlazo = "NO VARIARÁ"
    pcsub.impactoClasificacion = "MAYOR"
    pcsub.CambioAfectaAMaterial = "Material por entregar"
    svc.GuardarImpacto pcsub, db
    logs(0) = "1. Seeded full tecnica fields on idSol=" & idSol

    svc.PurgaTecnica idSol, db
    logs(1) = "2. PurgaTecnica called with injected db"

    Set rs = db.OpenRecordset( _
        "SELECT descripcionMaterialAfectado, descripcionPropuestaCambio, numPlanoEspecificacion, " & _
        "motivoMejorarSeguridad, incidenciaCoste, impactoClasificacion " & _
        "FROM tbDatosPCSUB WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If Nz(rs!descripcionMaterialAfectado, "") <> "" _
       Or Nz(rs!descripcionPropuestaCambio, "") <> "" _
       Or Nz(rs!numPlanoEspecificacion, "") <> "" _
       Or rs!motivoMejorarSeguridad <> False _
       Or Nz(rs!incidenciaCoste, "") <> "" _
       Or Nz(rs!impactoClasificacion, "") <> "" Then
        logs(2) = "3. FAILED: PurgaTecnica did not clear all tecnica fields"
        Test_PCSUB_BR010_PurgaTecnica_ClearsOnlyTecnicaFields = TestHelper.BuildJsonFail("tecnica fields not all cleared", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. All tecnica fields cleared (material, propuesta, motivos, impacto)"

    ' Parent solicitud must still exist (BR-010: "no mutar el resto de la fila/parentales")
    Set rs = db.OpenRecordset("SELECT idSolicitud FROM tbSolicitudes WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If rs.EOF Then
        logs(3) = "4. FAILED: parent tbSolicitudes row deleted by PurgaTecnica"
        Test_PCSUB_BR010_PurgaTecnica_ClearsOnlyTecnicaFields = TestHelper.BuildJsonFail("parent solicitud deleted", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. parent tbSolicitudes row preserved (BR-010 side-effect boundary honored)"

    logs(4) = "5. complete"
    logs(5) = "6. PASS"
    Test_PCSUB_BR010_PurgaTecnica_ClearsOnlyTecnicaFields = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR010_PurgaTecnica_ClearsOnlyTecnicaFields: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR010_PurgaTecnica_ClearsOnlyTecnicaFields = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR010_EliminarDictamenRAC_ClearsOnlyRACFields() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim rs As DAO.Recordset

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR010_EliminarDictamenRAC_ClearsOnlyRACFields = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR010_ID_DEL_RAC
    Call SeedMinimal(db, idSol)
    Dim pcsub As DatosPCSUB
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.racCodigo = "BR010-RAC-CODE"
    pcsub.racDecision = "APROBADO"
    pcsub.racNombre = "BR010-RAC-NAME"
    pcsub.observacionesRAC = "BR010-RAC-OBS"
    pcsub.observacionesRACDelegador = "BR010-DELEGADOR-OBS"
    pcsub.racNombreDelegador = "BR010-DELEGADOR-NAME"
    ' Set BOTH signer names so GuardarAprobacionSuministrador accepts the seed
    pcsub.firmaOficinaTecnicaSubSuministradorNombre = "BR010-APROB-OTICINA-PRESERVED"
    pcsub.firmaRepSubSuministradorNombre = "BR010-APROB-REPSUB-PRESERVED"
    svc.GuardarDictamenRAC pcsub, db
    svc.GuardarAprobacionSuministrador pcsub, db
    logs(0) = "1. Seeded RAC + Aprobacion on idSol=" & idSol

    svc.EliminarDictamenRAC idSol, db
    logs(1) = "2. EliminarDictamenRAC called with injected db"

    Set rs = db.OpenRecordset( _
        "SELECT racCodigo, racDecision, racNombre, observacionesRAC, observacionesRACDelegador, racNombreDelegador, " & _
        "firmaOficinaTecnicaSubSuministradorNombre FROM tbDatosPCSUB WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If Nz(rs!racCodigo, "") <> "" Or Nz(rs!racDecision, "") <> "" Or Nz(rs!racNombre, "") <> "" _
       Or Nz(rs!observacionesRAC, "") <> "" Or Nz(rs!observacionesRACDelegador, "") <> "" _
       Or Nz(rs!racNombreDelegador, "") <> "" Then
        logs(2) = "3. FAILED: RAC fields not all cleared"
        Test_PCSUB_BR010_EliminarDictamenRAC_ClearsOnlyRACFields = TestHelper.BuildJsonFail("RAC fields not cleared", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. All 6 RAC fields cleared (codigo+decision+nombre+obs+delegador*2)"

    If Nz(rs!firmaOficinaTecnicaSubSuministradorNombre, "") <> "BR010-APROB-OTICINA-PRESERVED" Then
        logs(3) = "4. FAILED: side-effect — Aprobacion signer was mutated by EliminarDictamenRAC"
        Test_PCSUB_BR010_EliminarDictamenRAC_ClearsOnlyRACFields = TestHelper.BuildJsonFail("Aprobacion signer mutated", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. Aprobacion signer PRESERVED (side-effect boundary honored)"

    logs(4) = "5. PASS"
    Test_PCSUB_BR010_EliminarDictamenRAC_ClearsOnlyRACFields = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR010_EliminarDictamenRAC_ClearsOnlyRACFields: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR010_EliminarDictamenRAC_ClearsOnlyRACFields = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR010_EliminarAprobacionSuministrador_ClearsOnlySignerFields() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim rs As DAO.Recordset

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR010_EliminarAprobacionSuministrador_ClearsOnlySignerFields = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR010_ID_DEL_APROB
    Call SeedMinimal(db, idSol)
    Dim pcsub As DatosPCSUB
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.firmaOficinaTecnicaSubSuministradorNombre = "BR010-APROB-OTICINA"
    pcsub.firmaRepSubSuministradorNombre = "BR010-APROB-REPSUB"
    pcsub.racCodigo = "BR010-RAC-PRESERVED"  ' control field
    pcsub.racDecision = "APROBADO"
    pcsub.racNombre = "BR010-RAC-NAME"
    svc.GuardarAprobacionSuministrador pcsub, db
    svc.GuardarDictamenRAC pcsub, db
    logs(0) = "1. Seeded Aprobacion + RAC on idSol=" & idSol

    svc.EliminarAprobacionSuministrador idSol, db
    logs(1) = "2. EliminarAprobacionSuministrador called with injected db"

    Set rs = db.OpenRecordset( _
        "SELECT firmaOficinaTecnicaSubSuministradorNombre, firmaRepSubSuministradorNombre, " & _
        "racCodigo, racDecision, racNombre FROM tbDatosPCSUB WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If Nz(rs!firmaOficinaTecnicaSubSuministradorNombre, "") <> "" _
       Or Nz(rs!firmaRepSubSuministradorNombre, "") <> "" Then
        logs(2) = "3. FAILED: signer fields not cleared"
        Test_PCSUB_BR010_EliminarAprobacionSuministrador_ClearsOnlySignerFields = TestHelper.BuildJsonFail("signer fields not cleared", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. Both signer fields cleared"

    If Nz(rs!racCodigo, "") <> "BR010-RAC-PRESERVED" Or Nz(rs!racDecision, "") <> "APROBADO" _
       Or Nz(rs!racNombre, "") <> "BR010-RAC-NAME" Then
        logs(3) = "4. FAILED: side-effect — RAC fields were mutated by EliminarAprobacionSuministrador"
        Test_PCSUB_BR010_EliminarAprobacionSuministrador_ClearsOnlySignerFields = TestHelper.BuildJsonFail("RAC fields mutated", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. RAC fields PRESERVED (side-effect boundary honored)"

    logs(4) = "5. PASS"
    Test_PCSUB_BR010_EliminarAprobacionSuministrador_ClearsOnlySignerFields = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR010_EliminarAprobacionSuministrador_ClearsOnlySignerFields: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR010_EliminarAprobacionSuministrador_ClearsOnlySignerFields = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR010_EliminarDecisionFinal_ClearsOnlyDecisionFields() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim idSol As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim rs As DAO.Recordset

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR010_EliminarDecisionFinal_ClearsOnlyDecisionFields = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR010_ID_DEL_DEC
    Call SeedMinimal(db, idSol)
    Dim pcsub As DatosPCSUB
    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.decisionFinal = "BR010-DECISION"
    pcsub.NombreFirmanteFinal = "BR010-FIRMANTE"
    pcsub.obsDecisionFinal = "BR010-OBS"
    pcsub.racCodigo = "BR010-RAC-PRESERVED"  ' control field
    pcsub.racDecision = "APROBADO"
    pcsub.racNombre = "BR010-RAC-NAME"
    svc.GuardarDecisionFinal pcsub, db
    svc.GuardarDictamenRAC pcsub, db
    logs(0) = "1. Seeded DecisionFinal + RAC on idSol=" & idSol

    svc.EliminarDecisionFinal idSol, db
    logs(1) = "2. EliminarDecisionFinal called with injected db"

    Set rs = db.OpenRecordset( _
        "SELECT decisionFinal, NombreFirmanteFinal, obsDecisionFinal, " & _
        "racCodigo, racDecision, racNombre FROM tbDatosPCSUB WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If Nz(rs!decisionFinal, "") <> "" Or Nz(rs!NombreFirmanteFinal, "") <> "" _
       Or Nz(rs!obsDecisionFinal, "") <> "" Then
        logs(2) = "3. FAILED: decision fields not cleared"
        Test_PCSUB_BR010_EliminarDecisionFinal_ClearsOnlyDecisionFields = TestHelper.BuildJsonFail("decision fields not cleared", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. All 3 decision fields cleared"

    If Nz(rs!racCodigo, "") <> "BR010-RAC-PRESERVED" Or Nz(rs!racDecision, "") <> "APROBADO" _
       Or Nz(rs!racNombre, "") <> "BR010-RAC-NAME" Then
        logs(3) = "4. FAILED: side-effect — RAC fields were mutated by EliminarDecisionFinal"
        Test_PCSUB_BR010_EliminarDecisionFinal_ClearsOnlyDecisionFields = TestHelper.BuildJsonFail("RAC fields mutated", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. RAC fields PRESERVED (side-effect boundary honored)"

    logs(4) = "5. PASS"
    Test_PCSUB_BR010_EliminarDecisionFinal_ClearsOnlyDecisionFields = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR010_EliminarDecisionFinal_ClearsOnlyDecisionFields: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR010_EliminarDecisionFinal_ClearsOnlyDecisionFields = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ============================================================================
' BR-011 — Injected DAO.Database preservation through service/repository
' (verifies the service does not bypass the injected db via getdb() / null
' the caller's db reference, preserving the form-layer contract that
' transitions must go through the service with the caller's db context)
' ============================================================================

Public Function Test_PCSUB_BR011_GuardarDictamenRAC_PreservesInjectedDb() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim setupError As String
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim db As DAO.Database
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim rs As DAO.Recordset

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR011_GuardarDictamenRAC_PreservesInjectedDb = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR011_ID_GUARDAR
    Call SeedMinimal(db, idSol)
    countBefore = CountPCSUBRows(db, idSol)
    If countBefore <> 1 Then
        logs(0) = "1. FAILED: expected 1 row, got " & countBefore
        Test_PCSUB_BR011_GuardarDictamenRAC_PreservesInjectedDb = TestHelper.BuildJsonFail("fixture setup", logs)
        GoTo CleanExit
    End If
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.racCodigo = "BR011-RAC"
    pcsub.racDecision = "APROBADO"
    pcsub.racNombre = "BR011-NAME"
    svc.GuardarDictamenRAC pcsub, db
    logs(1) = "2. GuardarDictamenRAC called with injected db"

    ' BR-011 contract: caller's db reference must remain usable after the call
    If db Is Nothing Then
        logs(2) = "3. FAILED: injected db reference was cleared (BR-011 violation)"
        Test_PCSUB_BR011_GuardarDictamenRAC_PreservesInjectedDb = TestHelper.BuildJsonFail("injected db reference cleared", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. injected db reference remained usable after GuardarDictamenRAC"

    countAfter = CountPCSUBRows(db, idSol)
    If countAfter <> 1 Then
        logs(3) = "4. FAILED: row count changed"
        Test_PCSUB_BR011_GuardarDictamenRAC_PreservesInjectedDb = TestHelper.BuildJsonFail("row count changed", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=" & countAfter & " (expected 1)"

    ' Verify the write went through the injected db (not via getdb() fallback)
    Set rs = db.OpenRecordset("SELECT racCodigo, racDecision FROM tbDatosPCSUB WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If Nz(rs!racCodigo, "") <> "BR011-RAC" Or Nz(rs!racDecision, "") <> "APROBADO" Then
        logs(4) = "5. FAILED: RAC fields not persisted through injected db"
        Test_PCSUB_BR011_GuardarDictamenRAC_PreservesInjectedDb = TestHelper.BuildJsonFail("RAC not persisted through injected db", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. RAC fields persisted and re-read through the same injected db"

    logs(5) = "6. PASS"
    Test_PCSUB_BR011_GuardarDictamenRAC_PreservesInjectedDb = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR011_GuardarDictamenRAC_PreservesInjectedDb: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR011_GuardarDictamenRAC_PreservesInjectedDb = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_BR011_ActualizarAprobacionSuministrador_PreservesInjectedDb() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim setupError As String
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Call SetupSandbox(setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_PCSUB_BR011_ActualizarAprobacionSuministrador_PreservesInjectedDb = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If
    Set db = TestHelper.GetTestDb()
    Call TeardownAll(db)

    idSol = BR011_ID_ACTUALIZAR
    Call SeedMinimal(db, idSol)
    countBefore = CountPCSUBRows(db, idSol)
    If countBefore <> 1 Then
        logs(0) = "1. FAILED: expected 1 row, got " & countBefore
        Test_PCSUB_BR011_ActualizarAprobacionSuministrador_PreservesInjectedDb = TestHelper.BuildJsonFail("fixture setup", logs)
        GoTo CleanExit
    End If
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    Dim pcsub As New DatosPCSUB
    pcsub.idSolicitud = idSol
    pcsub.firmaOficinaTecnicaSubSuministradorNombre = "BR011-OTICINA"
    pcsub.firmaRepSubSuministradorNombre = "BR011-REPSUB"
    DatosPCSUBRepositorio.ActualizarAprobacionSuministrador pcsub, db
    logs(1) = "2. Repository called with injected db"

    ' BR-011 contract (Issue #7 regression guard)
    If db Is Nothing Then
        logs(2) = "3. FAILED: injected db reference was cleared (BR-011 violation)"
        Test_PCSUB_BR011_ActualizarAprobacionSuministrador_PreservesInjectedDb = TestHelper.BuildJsonFail("injected db reference cleared", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. injected db reference remained usable after ActualizarAprobacionSuministrador"

    countAfter = CountPCSUBRows(db, idSol)
    If countAfter <> 1 Then
        logs(3) = "4. FAILED: row count changed"
        Test_PCSUB_BR011_ActualizarAprobacionSuministrador_PreservesInjectedDb = TestHelper.BuildJsonFail("row count changed", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=" & countAfter & " (expected 1)"

    Set rs = db.OpenRecordset( _
        "SELECT firmaOficinaTecnicaSubSuministradorNombre, firmaRepSubSuministradorNombre " & _
        "FROM tbDatosPCSUB WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If Nz(rs!firmaOficinaTecnicaSubSuministradorNombre, "") <> "BR011-OTICINA" _
       Or Nz(rs!firmaRepSubSuministradorNombre, "") <> "BR011-REPSUB" Then
        logs(4) = "5. FAILED: signer fields not persisted through injected db"
        Test_PCSUB_BR011_ActualizarAprobacionSuministrador_PreservesInjectedDb = TestHelper.BuildJsonFail("signer fields not persisted", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. signer fields persisted and re-read through same injected db"

    logs(5) = "6. PASS"
    Test_PCSUB_BR011_ActualizarAprobacionSuministrador_PreservesInjectedDb = TestHelper.BuildJsonOk("true", logs)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not db Is Nothing Then TeardownAll db
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_BR011_ActualizarAprobacionSuministrador_PreservesInjectedDb: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_BR011_ActualizarAprobacionSuministrador_PreservesInjectedDb = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function
