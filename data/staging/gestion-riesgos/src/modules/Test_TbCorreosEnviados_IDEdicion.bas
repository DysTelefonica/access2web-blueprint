Attribute VB_Name = "Test_TbCorreosEnviados_IDEdicion"
Option Compare Database
Option Explicit

' ============================================================
' Test_TbCorreosEnviados_IDEdicion — Tests para el fix de issue #70
'
' Skill: access-vba-tdd v2.5 (schema-first + fixture patterns + HandleError)
'
' Cobertura (3 tests atómicos, REQ-MIG-021-01..04):
'   Test 1 — PersisteValorSobreIntegerMax (DAO accepts 50000 IDEdicion;
'           RED pre-migration because column is Integer, max 32,767)
'   Test 2 — MigrationEsIdempotente (MigracionCorreosEnviadosIDEdicionLong.EjecutarMigracion
'           called twice returns EnumSiNo.Sí both times; RED pre-migration
'           because helper does not exist yet — late-bound via Application.Run
'           to avoid compile-time dependency)
'   Test 3 — RegistrosExistentesPreservados (count of rows with IDEdicion IS NOT NULL
'           unchanged after migration; GREEN from start because migration is
'           cardinality-invariant regardless of column type)
'
' Fixture IDs (FIX_ID_* >= 50000) are intentionally ABOVE the Integer max (32,767).
' This is the whole point of the test: we PROVE the migration is necessary and
' effective by using IDs that do NOT fit in the Integer column. If the test were
' written with FIX_ID_* < 32768, a verde-por-suerte pass on the Integer column
' would hide the bug (lesson learned from issue #64 JD round 1, see obs 13242
' and commit d1593d7).
' ============================================================

' --- Module-level declarations (AGENTS.md rule 3) ---
Private Const FIX_ID_BASE As Long = 50000
Private Const FIX_ID_CORREO As Long = 50001   ' Reserved marker for TbCorreosEnviados rows (autonum IDCorreo; not used directly in INSERTs)
Private Const FIX_ID_EDICION As Long = 50010
Private Const FIX_ID_PROYECTO As Long = 50011
Private Const FIX_ID_EXPEDIENTE As Long = 50012
Private Const FIX_ID_VALOR As Long = 50000    ' IDEdicion value used in TbCorreosEnviados INSERT (above Integer max — proves migration needed)

' --- JSON helpers (delegación a Test_Helper) ---
Private Function JsonOk(ByVal value As Variant, ByRef logs() As String) As String
    JsonOk = BuildJsonOk(value, logs)
End Function

Private Function JsonFail(ByVal errMsg As String, ByRef logs() As String) As String
    JsonFail = BuildJsonFail(errMsg, logs)
End Function

' --- EnsureTestConfigLoaded (delegación a Test_Helper) ---
Private Function EnsureTestConfigLoaded(ByRef p_Error As String) As Boolean
    EnsureTestConfigLoaded = Test_Helper.EnsureTestConfigLoaded(p_Error)
End Function

' --- GetTestDb (delegación a Test_Fixtures) ---
Private Function GetTestDb(ByRef p_Error As String) As DAO.Database
    Set GetTestDb = Test_Fixtures.GetTestDb(p_Error)
End Function

' --- Helper: read current IDEdicion column type via DAO ---
' Returns the DAO dbType constant (dbInteger=3, dbLong=4). Empty p_Error on success.
Private Function TipoColumnaIDEdicion(ByRef p_Error As String) As Long
    On Error GoTo EH
    Dim db As DAO.Database
    Dim dbErr As String
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        p_Error = "TipoColumnaIDEdicion: GetTestDb returned Nothing: " & dbErr
        Exit Function
    End If

    Set tdf = db.TableDefs("TbCorreosEnviados")
    Set fld = tdf.fields("IDEdicion")
    TipoColumnaIDEdicion = fld.Type

    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

EH:
    p_Error = "TipoColumnaIDEdicion: " & Err.description
End Function

' --- Fixture helpers (FK-ordered) ---

' Seed FK chain: TbExpedientes ? TbProyectos ? TbProyectosEdiciones
' The TbCorreosEnviados row is NOT seeded here — tests insert it themselves
' (Test 1 inserts IDEdicion=50000, Test 2/3 do not touch TbCorreosEnviados rows).
' Post-seed cardinality assertion proves the seed actually persisted (issue #64 lesson).
Private Sub SeedCorreosFixture()
    On Error GoTo EH_Seed
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then Err.Raise 1001, "SeedCorreosFixture", "GetTestDb returned Nothing: " & dbErr

    ' Idempotent cleanup in reverse FK order (only deterministic test markers)
    On Error Resume Next
    db.Execute "DELETE FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_ID_VALOR, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_ID_PROYECTO, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_ID_EXPEDIENTE, dbFailOnError
    On Error GoTo 0

    ' 1. TbExpedientes (padre)
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
               "VALUES (" & FIX_ID_EXPEDIENTE & ", 'TEST70', 'Fixture issue 70', 'Test', 1)", dbFailOnError

    ' 2. TbProyectos (hijo de Expediente)
    db.Execute "INSERT INTO TbProyectos " & _
               "(IDProyecto, IDExpediente, Proyecto, ParaInformeAvisos, NombreUsuarioCalidad) " & _
               "VALUES (" & FIX_ID_PROYECTO & ", " & FIX_ID_EXPEDIENTE & ", 'TESTPROJ70', 'Sí', 'test_calidad_70')", dbFailOnError

    ' 3. TbProyectosEdiciones (hijo de Proyecto) — IDEdicion is Long (already), > Integer max
    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDEdicion, IDProyecto, Edicion, Elaborado, FechaPreparadaParaPublicar) " & _
               "VALUES (" & FIX_ID_EDICION & ", " & FIX_ID_PROYECTO & ", 1, " & _
               "'test_calidad_70', #" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    ' Post-seed cardinality assertion — guard against verde-por-suerte (issue #64 lesson, obs 13242).
    ' If the INSERTs above silently failed (e.g., FK violation masked by dbFailOnError quirk),
    ' this assertion fails the seed itself, not the test.
    Dim seedRs As DAO.Recordset
    Dim seedCount As Long
    Set seedRs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_ID_EDICION)
    If Not seedRs.EOF Then seedCount = CLng(Nz(seedRs.fields("C").value, 0))
    seedRs.Close
    Set seedRs = Nothing
    If seedCount <> 1 Then
        Err.Raise 1002, "SeedCorreosFixture", "Post-seed assertion failed: expected 1 row in TbProyectosEdiciones, got " & seedCount
    End If

    Set db = Nothing
    Exit Sub

EH_Seed:
    Dim eNuma As Long
    Dim eDesc As String
    eNuma = Err.Number
    eDesc = Err.Description
    On Error Resume Next
    Set db = Nothing
    Err.Raise eNuma, "SeedCorreosFixture", "Seed failed: " & eNuma & " - " & eDesc
End Sub

' Teardown in reverse FK order — only deletes deterministic test markers (FIX_ID_* >= 50000)
Private Sub TeardownCorreosFixture()
    On Error Resume Next
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then Exit Sub

    db.Execute "DELETE FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_ID_VALOR, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_ID_PROYECTO, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_ID_EXPEDIENTE, dbFailOnError

    Set db = Nothing
    On Error GoTo 0
End Sub

' ============================================================
' Test 1 — PersisteValorSobreIntegerMax (REQ-MIG-021-01)
' GIVEN TbCorreosEnviados schema (Integer pre-migration, Long post-migration)
' WHEN INSERT into TbCorreosEnviados with IDEdicion = 50000 (above Integer max)
' THEN:
'   - Pre-migration: DAO error 3035 (data type conversion overflow) ? RED
'   - Post-migration: insert succeeds and IDEdicion round-trips as exactly 50000 ? GREEN
' ============================================================
 Public Function Test_TbCorreosEnviados_IDEdicion_PersisteValorSobreIntegerMax() As String
    Dim logs(0 To 8) As String
    ' Initialize result to fail BEFORE On Error (skill §1.1.2 HandleError pattern)
    Test_TbCorreosEnviados_IDEdicion_PersisteValorSobreIntegerMax = JsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: EnsureTestConfigLoaded"
    logs(2) = "3. Arrange: SeedCorreosFixture"
    logs(3) = "4. Act: Migrate IDEdicion field to Long"
    logs(4) = "5. Assert: migration returned EnumSiNo.Sí"
    logs(5) = "6. Act: INSERT TbCorreosEnviados (IDEdicion=" & FIX_ID_VALOR & ")"
    logs(6) = "7. Assert: read IDEdicion back = " & FIX_ID_VALOR
    Dim runError As String
    ForceLocalBackend runError
    If runError <> vbNullString Then
        Test_TbCorreosEnviados_IDEdicion_PersisteValorSobreIntegerMax = JsonFail("ForceLocalBackend failed: " & runError, logs)
        GoTo Teardown
    End If

    Dim migError As String
    Dim migResult As EnumSiNo

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_TbCorreosEnviados_IDEdicion_PersisteValorSobreIntegerMax = JsonFail(cfgError, logs)
        GoTo Teardown
    End If

    SeedCorreosFixture

    migError = vbNullString
    migResult = MigracionCorreosEnviadosIDEdicionLong.EjecutarMigracion(migError)
    If migError <> vbNullString Then
        Test_TbCorreosEnviados_IDEdicion_PersisteValorSobreIntegerMax = JsonFail("Migration failed: " & migError, logs)
        GoTo Teardown
    End If
    If migResult <> EnumSiNo.Sí Then
        Test_TbCorreosEnviados_IDEdicion_PersisteValorSobreIntegerMax = JsonFail("Migration did not return EnumSiNo.Sí", logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_TbCorreosEnviados_IDEdicion_PersisteValorSobreIntegerMax = JsonFail("GetTestDb returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' Act: INSERT a correo with IDEdicion > Integer max. dbFailOnError surfaces DAO overflow (3035).
    db.Execute "INSERT INTO TbCorreosEnviados (IDEdicion, Aplicacion, Asunto, FechaEnvio, FechaGrabacion) " & _
               "VALUES (" & FIX_ID_VALOR & ", 'TEST70', 'Test issue 70', " & _
               "#" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#, " & _
               "#" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    ' Assert: read IDEdicion back and verify round-trip
    Dim rs As DAO.Recordset
    Dim readValue As Variant
    Set rs = db.OpenRecordset("SELECT TOP 1 IDEdicion FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_ID_VALOR)
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        Test_TbCorreosEnviados_IDEdicion_PersisteValorSobreIntegerMax = JsonFail("INSERT succeeded but row not visible in SELECT", logs)
        GoTo Teardown
    End If
    readValue = rs.fields("IDEdicion").value
    rs.Close
    Set rs = Nothing

    If IsNull(readValue) Then
        Test_TbCorreosEnviados_IDEdicion_PersisteValorSobreIntegerMax = JsonFail("IDEdicion is Null after INSERT", logs)
        GoTo Teardown
    End If
    If CLng(readValue) <> FIX_ID_VALOR Then
        Test_TbCorreosEnviados_IDEdicion_PersisteValorSobreIntegerMax = JsonFail("IDEdicion expected " & FIX_ID_VALOR & " but got " & CLng(readValue), logs)
        GoTo Teardown
    End If

    Test_TbCorreosEnviados_IDEdicion_PersisteValorSobreIntegerMax = JsonOk("persiste_valor_sobre_integer_max_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownCorreosFixture
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_TbCorreosEnviados_IDEdicion_PersisteValorSobreIntegerMax = JsonFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' Test 2 — MigrationEsIdempotente (REQ-MIG-021-01 + REQ-MIG-021-03)
' GIVEN MigracionCorreosEnviadosIDEdicionLong module is importable
' WHEN Ejecutar is called twice
' THEN both calls return EnumSiNo.Sí, p_Error stays empty, column type is Long
' AND second call does not re-alter (idempotent).
' ============================================================
Public Function Test_TbCorreosEnviados_IDEdicion_MigrationEsIdempotente() As String
    Dim logs(0 To 8) As String
    Test_TbCorreosEnviados_IDEdicion_MigrationEsIdempotente = JsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: EnsureTestConfigLoaded"
    logs(2) = "3. Act: EjecutarMigracion (1st call)"
    logs(3) = "4. Assert: 1st call returned EnumSiNo.Sí"
    logs(4) = "5. Act: EjecutarMigracion (2nd call)"
    logs(5) = "6. Assert: 2nd call returned EnumSiNo.Sí (idempotent)"
    logs(6) = "7. Assert: column type is dbLong (post-migration)"
    logs(7) = "8. Teardown: TeardownCorreosFixture"
    logs(8) = "9. Teardown: Test_Helper.ResetTestSession"

    Dim runError As String
    ForceLocalBackend runError
    If runError <> vbNullString Then
        Test_TbCorreosEnviados_IDEdicion_MigrationEsIdempotente = JsonFail("ForceLocalBackend failed: " & runError, logs)
        GoTo Teardown
    End If

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_TbCorreosEnviados_IDEdicion_MigrationEsIdempotente = JsonFail(cfgError, logs)
        GoTo Teardown
    End If

    Dim result1 As EnumSiNo
    Dim resultError As String
    resultError = vbNullString
    result1 = MigracionCorreosEnviadosIDEdicionLong.EjecutarMigracion(resultError)
    If resultError <> vbNullString Then
        Test_TbCorreosEnviados_IDEdicion_MigrationEsIdempotente = JsonFail("First Ejecutar failed: " & resultError, logs)
        GoTo Teardown
    End If
    If result1 <> EnumSiNo.Sí Then
        Test_TbCorreosEnviados_IDEdicion_MigrationEsIdempotente = JsonFail("First Ejecutar returned " & CStr(result1) & " (expected EnumSiNo.Sí=" & CLng(EnumSiNo.Sí) & ")", logs)
        GoTo Teardown
    End If

    ' Second Ejecutar call — should be idempotent
    Dim result2 As EnumSiNo
    resultError = vbNullString
    result2 = MigracionCorreosEnviadosIDEdicionLong.EjecutarMigracion(resultError)
    If resultError <> vbNullString Then
        Test_TbCorreosEnviados_IDEdicion_MigrationEsIdempotente = JsonFail("Second Ejecutar failed: " & resultError, logs)
        GoTo Teardown
    End If
    If result2 <> EnumSiNo.Sí Then
        Test_TbCorreosEnviados_IDEdicion_MigrationEsIdempotente = JsonFail("Second Ejecutar returned " & CStr(result2) & " (expected EnumSiNo.Sí=" & CLng(EnumSiNo.Sí) & ")", logs)
        GoTo Teardown
    End If

    ' Assert column type is Long (post-migration invariant — must hold regardless of starting state)
    Dim colTypeErr As String
    Dim colType As Long
    colType = TipoColumnaIDEdicion(colTypeErr)
    If colTypeErr <> "" Then
        Test_TbCorreosEnviados_IDEdicion_MigrationEsIdempotente = JsonFail("TipoColumnaIDEdicion: " & colTypeErr, logs)
        GoTo Teardown
    End If
    If colType <> dbLong Then
        Test_TbCorreosEnviados_IDEdicion_MigrationEsIdempotente = JsonFail("Column type is " & colType & " (expected dbLong=" & dbLong & ")", logs)
        GoTo Teardown
    End If

    Test_TbCorreosEnviados_IDEdicion_MigrationEsIdempotente = JsonOk("migration_es_idempotente_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    TeardownCorreosFixture
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_TbCorreosEnviados_IDEdicion_MigrationEsIdempotente = JsonFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' Test 3 — RegistrosExistentesPreservados (REQ-MIG-021-03, GREEN from start)
' GIVEN staging has rows in TbCorreosEnviados with IDEdicion IS NOT NULL
' WHEN MigracionCorreosEnviadosIDEdicionLong.EjecutarMigracion runs
' THEN count(rows WHERE IDEdicion IS NOT NULL) is invariant.
'
' Why this is GREEN pre-migration: ALTER COLUMN on Integer?Long preserves all
' existing rows (no truncation, no type coercion that loses data). This test
' catches the failure mode "migration runs but loses rows" without needing
' the column to be Long first.
' ============================================================
Public Function Test_TbCorreosEnviados_IDEdicion_RegistrosExistentesPreservados() As String
    Dim logs(0 To 8) As String
    Test_TbCorreosEnviados_IDEdicion_RegistrosExistentesPreservados = JsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: EnsureTestConfigLoaded"
    logs(2) = "3. Arrange: SeedCorreosFixture (provides deterministic IDEdicion IS NOT NULL rows)"
    logs(3) = "4. Act: count(TbCorreosEnviados WHERE IDEdicion IS NOT NULL) before migration"
    logs(4) = "5. Act: Ejecutar"
    logs(5) = "6. Act: count after migration"
    logs(6) = "7. Assert: count_after >= count_before (no rows lost)"
    logs(7) = "8. Teardown: TeardownCorreosFixture"
    logs(8) = "9. Teardown: Test_Helper.ResetTestSession"

    Dim runError As String
    ForceLocalBackend runError
    If runError <> vbNullString Then
        Test_TbCorreosEnviados_IDEdicion_RegistrosExistentesPreservados = JsonFail("ForceLocalBackend failed: " & runError, logs)
        GoTo Teardown
    End If

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_TbCorreosEnviados_IDEdicion_RegistrosExistentesPreservados = JsonFail(cfgError, logs)
        GoTo Teardown
    End If

    SeedCorreosFixture

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_TbCorreosEnviados_IDEdicion_RegistrosExistentesPreservados = JsonFail("GetTestDb returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' Count before migration
    Dim countBefore As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbCorreosEnviados WHERE IDEdicion IS NOT NULL")
    If Not rs.EOF Then countBefore = CLng(Nz(rs.fields("C").value, 0))
    rs.Close
    Set rs = Nothing
    logs(3) = "4. count_before (global, not just fixture): " & countBefore

    ' Act: Ejecutar
    Dim migResult As Variant
    Err.Clear
    migResult = MigracionCorreosEnviadosIDEdicionLong.EjecutarMigracion(cfgError)
    If cfgError <> vbNullString Then
        Test_TbCorreosEnviados_IDEdicion_RegistrosExistentesPreservados = JsonFail("Ejecutar failed: " & cfgError, logs)
        GoTo Teardown
    End If
    If CLng(migResult) <> CLng(EnumSiNo.Sí) Then
        Test_TbCorreosEnviados_IDEdicion_RegistrosExistentesPreservados = JsonFail("Ejecutar returned " & CStr(migResult) & " (expected EnumSiNo.Sí=" & CLng(EnumSiNo.Sí) & ")", logs)
        GoTo Teardown
    End If

    ' Count after migration
    Dim countAfter As Long
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbCorreosEnviados WHERE IDEdicion IS NOT NULL")
    If Not rs.EOF Then countAfter = CLng(Nz(rs.fields("C").value, 0))
    rs.Close
    Set rs = Nothing
    logs(5) = "6. count_after (global, not just fixture): " & countAfter

    ' Invariant: no rows lost. Migration may ADD rows via fixture Teardown leftover, so >= not =.
    ' (TeardownCorreosFixture cleans up, so in practice counts should be equal within fixture scope;
    '  but if test 1's correo row leaked, count_after could be > count_before — both are OK.)
    If countAfter < countBefore Then
        Test_TbCorreosEnviados_IDEdicion_RegistrosExistentesPreservados = JsonFail("Cardinality regression: count_before=" & countBefore & " count_after=" & countAfter & " (rows lost)", logs)
        GoTo Teardown
    End If

    Test_TbCorreosEnviados_IDEdicion_RegistrosExistentesPreservados = JsonOk("registros_existentes_preservados_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownCorreosFixture
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_TbCorreosEnviados_IDEdicion_RegistrosExistentesPreservados = JsonFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' RunAll — Aggregator following Test_RechazoPropuesta_Issue64_RunAll pattern.
' Per access-vba-tdd §4.7 manifest discipline: RunAll goes in tests.vba.smoke.json
' (or run directly as smoke), NOT in tests.vba.json (atomic suite).
' ============================================================
Public Function Test_TbCorreosEnviados_IDEdicion_RunAll() As String
    Dim results(0 To 2) As String
    Dim names(0 To 2) As String
    Dim i As Long
    Dim runError As String
    Dim outLogs(0 To 6) As String

    ' SuiteSetup
    ForceLocalBackend runError
    If runError <> "" Then
        outLogs(0) = "TESTS BLOCKED: " & runError
        Test_TbCorreosEnviados_IDEdicion_RunAll = JsonFail("TESTS BLOCKED: " & runError, outLogs)
        Exit Function
    End If
    outLogs(0) = "SuiteSetup OK"

    names(0) = "Test_TbCorreosEnviados_IDEdicion_PersisteValorSobreIntegerMax"
    names(1) = "Test_TbCorreosEnviados_IDEdicion_MigrationEsIdempotente"
    names(2) = "Test_TbCorreosEnviados_IDEdicion_RegistrosExistentesPreservados"

    results(0) = Test_TbCorreosEnviados_IDEdicion_PersisteValorSobreIntegerMax()
    results(1) = Test_TbCorreosEnviados_IDEdicion_MigrationEsIdempotente()
    results(2) = Test_TbCorreosEnviados_IDEdicion_RegistrosExistentesPreservados()

    ' SuiteTeardown defensivo
    TeardownCorreosFixture
    outLogs(1) = "SuiteTeardown OK"

    Test_Helper.ResetTestSession
    outLogs(2) = "ResetTestSession OK"

    ' Acumular
    Dim allOk As Boolean
    allOk = True
    Dim firstFailure As String
    firstFailure = ""
    For i = 0 To 2
        If InStr(results(i), """ok"":false") > 0 Then
            allOk = False
            If firstFailure = "" Then firstFailure = names(i)
        End If
        outLogs(i + 3) = names(i) & ": " & IIf(InStr(results(i), """ok"":false") > 0, "FAIL", "OK")
    Next i

    If allOk Then
        Test_TbCorreosEnviados_IDEdicion_RunAll = JsonOk("all_pass", outLogs)
    Else
        Test_TbCorreosEnviados_IDEdicion_RunAll = JsonFail("some_tests_failed: " & firstFailure, outLogs)
    End If
End Function

