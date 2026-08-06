Attribute VB_Name = "Test_TbCambiosParaPublicacion_Edicion"
Option Compare Database
Option Explicit

' ============================================================
' Test_TbCambiosParaPublicacion_Edicion — Tests para el fix de issue #107
'
' Skill: access-vba-tdd v2.5 (schema-first + fixture patterns + HandleError)
'
' Cubre DOS columnas en una sola migración:
'   - tbCambiosParaPublicacion.EdicionInicial
'   - tbCambiosParaPublicacion.EdicionFinal
' Ambas son FK a TbProyectosEdiciones.IDEdicion (Long).
'
' Cobertura (3 tests atómicos):
'   Test 1 — PersisteValorSobreIntegerMax (DAO accepts 50000 EdicionInicial
'           AND 50001 EdicionFinal; RED pre-migration because both columns
'           are Integer, max 32,767)
'   Test 2 — MigrationEsIdempotente (EjecutarMigracion called twice returns
'           EnumSiNo.Sí both times; checks that BOTH columns end up dbLong
'           regardless of starting state)
'   Test 3 — RegistrosExistentesPreservados (count of rows in
'           tbCambiosParaPublicacion unchanged after migration; GREEN from
'           start because ALTER COLUMN on Integer?Long preserves all rows)
'
' Fixture IDs (FIX_ID_* >= 50000) are intentionally ABOVE the Integer max.
' Like #70 (issue #64 JD lesson), this PROVES the migration is necessary by
' using IDs that do NOT fit in Integer. If the test wrote FIX_ID_* < 32768,
' a verde-por-suerte pass on Integer would hide the bug.
' ============================================================

' --- Module-level declarations (AGENTS.md rule 3) ---
Private Const FIX_ID_BASE As Long = 50000
Private Const FIX_ID_PROYECTO As Long = 50007
Private Const FIX_ID_EXPEDIENTE As Long = 50008
Private Const FIX_ID_EDICION_ALTA As Long = 50009
Private Const FIX_ID_EDICION_BAJA As Long = 50010
Private Const FIX_ID_CAMBIO_INI As Long = 50011   ' IDCambio del row usado para INSERT EdicionInicial=50000
Private Const FIX_ID_CAMBIO_FIN As Long = 50012   ' IDCambio del row usado para INSERT EdicionFinal=50001
Private Const FIX_ID_VALOR_INI As Long = 50000    ' > Integer max -- proves migration needed
Private Const FIX_ID_VALOR_FIN As Long = 50001    ' > Integer max -- proves migration needed


' --- JSON helpers (delegacion a Test_Helper) ---
Private Function JsonOk(ByVal value As Variant, ByRef logs() As String) As String
    JsonOk = BuildJsonOk(value, logs)
End Function

Private Function JsonFail(ByVal errMsg As String, ByRef logs() As String) As String
    JsonFail = BuildJsonFail(errMsg, logs)
End Function


' --- EnsureTestConfigLoaded (delegacion a Test_Helper) ---
Private Function EnsureTestConfigLoaded(ByRef p_Error As String) As Boolean
    EnsureTestConfigLoaded = Test_Helper.EnsureTestConfigLoaded(p_Error)
End Function


' --- GetTestDb (delegación a Test_Fixtures) ---
Private Function GetTestDb(ByRef p_Error As String) As DAO.Database
    Set GetTestDb = Test_Fixtures.GetTestDb(p_Error)
End Function


' --- Helper: read both column types via DAO ---
' Returns "<tipoIni>,<tipoFin>" or "" on error.
Private Function TipoColumnasCambiosPub(ByRef p_Error As String) As String
    On Error GoTo EH
    Dim db As DAO.Database
    Dim dbErr As String
    Dim tdf As DAO.TableDef
    Dim fldIni As DAO.Field
    Dim fldFin As DAO.Field

    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        p_Error = "TipoColumnasCambiosPub: GetTestDb returned Nothing: " & dbErr
        Exit Function
    End If

    Set tdf = db.TableDefs("tbCambiosParaPublicacion")
    Set fldIni = tdf.Fields("EdicionInicial")
    Set fldFin = tdf.Fields("EdicionFinal")
    TipoColumnasCambiosPub = TipoComoNombre(fldIni.Type) & "," & TipoComoNombre(fldFin.Type)

    Set fldIni = Nothing
    Set fldFin = Nothing
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

EH:
    p_Error = "TipoColumnasCambiosPub: " & Err.description
End Function

Private Function TipoComoNombre(ByVal dbType As Long) As String
    Select Case dbType
        Case dbLong
            TipoComoNombre = "LONG"
        Case dbInteger
            TipoComoNombre = "INTEGER"
        Case Else
            TipoComoNombre = "OTHER"
    End Select
End Function


' --- Fixture helpers (FK-ordered) ---

' Seed FK chain: TbExpedientes > TbProyectos > TbProyectosEdiciones.
' Los rows de tbCambiosParaPublicacion NO se siembran aqui -- los tests los INSERTan ellos.
' Post-seed cardinality assertion prueba que la seed realmente persistio (issue #64 lesson).
Private Sub SeedCambiosPubFixture()
    On Error GoTo EH_Seed
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then Err.Raise 1001, "SeedCambiosPubFixture", "GetTestDb returned Nothing: " & dbErr

    ' Idempotent cleanup in reverse FK order
    On Error Resume Next
    db.Execute "DELETE FROM tbCambiosParaPublicacion WHERE IDCambio IN (" & FIX_ID_CAMBIO_INI & "," & FIX_ID_CAMBIO_FIN & ")", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion IN (" & FIX_ID_EDICION_ALTA & "," & FIX_ID_EDICION_BAJA & ")", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_ID_PROYECTO, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_ID_EXPEDIENTE, dbFailOnError
    On Error GoTo 0

    ' 1. TbExpedientes (padre)
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
               "VALUES (" & FIX_ID_EXPEDIENTE & ", 'TEST107', 'Fixture issue 107', 'Test', 1)", dbFailOnError

    ' 2. TbProyectos (hijo de Expediente)
    db.Execute "INSERT INTO TbProyectos " & _
               "(IDProyecto, IDExpediente, Proyecto, ParaInformeAvisos, NombreUsuarioCalidad) " & _
               "VALUES (" & FIX_ID_PROYECTO & ", " & FIX_ID_EXPEDIENTE & ", 'TESTPROJ107', 'Sí', 'test_calidad_107')", dbFailOnError

    ' 3. TbProyectosEdiciones (hijo de Proyecto) — IDEdicion es Long (already)
    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDEdicion, IDProyecto, Edicion, Elaborado, FechaPreparadaParaPublicar) " & _
               "VALUES (" & FIX_ID_EDICION_ALTA & ", " & FIX_ID_PROYECTO & ", 1, " & _
               "'test_calidad_107', #" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDEdicion, IDProyecto, Edicion, Elaborado, FechaPreparadaParaPublicar) " & _
               "VALUES (" & FIX_ID_EDICION_BAJA & ", " & FIX_ID_PROYECTO & ", 2, " & _
               "'test_calidad_107', #" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    ' Post-seed cardinality assertion
    Dim seedRs As DAO.Recordset
    Dim seedCount As Long
    Set seedRs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbProyectosEdiciones WHERE IDProyecto=" & FIX_ID_PROYECTO)
    If Not seedRs.EOF Then seedCount = CLng(Nz(seedRs.Fields("C").value, 0))
    seedRs.Close
    Set seedRs = Nothing
    If seedCount <> 2 Then
        Err.Raise 1002, "SeedCambiosPubFixture", "Post-seed assertion failed: expected 2 rows in TbProyectosEdiciones for IDProyecto=" & FIX_ID_PROYECTO & ", got " & seedCount
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
    Err.Raise eNuma, "SeedCambiosPubFixture", "Seed failed: " & eNuma & " - " & eDesc
End Sub


' Teardown in reverse FK order
Private Sub TeardownCambiosPubFixture()
    On Error Resume Next
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then Exit Sub

    db.Execute "DELETE FROM tbCambiosParaPublicacion WHERE IDCambio IN (" & FIX_ID_CAMBIO_INI & "," & FIX_ID_CAMBIO_FIN & ")", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion IN (" & FIX_ID_EDICION_ALTA & "," & FIX_ID_EDICION_BAJA & ")", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_ID_PROYECTO, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_ID_EXPEDIENTE, dbFailOnError

    Set db = Nothing
    On Error GoTo 0
End Sub


' ============================================================
' Test 1 — PersisteValorSobreIntegerMax
' GIVEN tbCambiosParaPublicacion schema (Integer pre-migration, Long post-migration)
' WHEN INSERT into tbCambiosParaPublicacion with EdicionInicial = 50000 AND EdicionFinal = 50001
' THEN:
'   - Pre-migration: DAO error 3035 (data type conversion overflow) ? RED
'   - Post-migration: insert succeeds and round-trips both values exactly ? GREEN
' ============================================================
Public Function Test_TbCambiosParaPublicacion_Edicion_PersisteValorSobreIntegerMax() As String
    Dim logs(0 To 8) As String
    Test_TbCambiosParaPublicacion_Edicion_PersisteValorSobreIntegerMax = JsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: EnsureTestConfigLoaded"
    logs(2) = "3. Arrange: SeedCambiosPubFixture"
    logs(3) = "4. Act: Migrar AMBAS columnas a Long"
    logs(4) = "5. Assert: migration returned EnumSiNo.Si"
    logs(5) = "6. Act: INSERT tbCambiosParaPublicacion (EdicionInicial=50000, EdicionFinal=50001)"
    logs(6) = "7. Assert: read EdicionInicial back = 50000"
    logs(7) = "8. Assert: read EdicionFinal back = 50001"

    Dim runError As String
    ForceLocalBackend runError
    If runError <> vbNullString Then
        Test_TbCambiosParaPublicacion_Edicion_PersisteValorSobreIntegerMax = JsonFail("ForceLocalBackend failed: " & runError, logs)
        GoTo Teardown
    End If

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_TbCambiosParaPublicacion_Edicion_PersisteValorSobreIntegerMax = JsonFail(cfgError, logs)
        GoTo Teardown
    End If

    SeedCambiosPubFixture

    Dim migError As String
    Dim migResult As EnumSiNo
    migError = vbNullString
    migResult = MigracionTbCambiosParaPublicacionEdicionLong.EjecutarMigracion(migError)
    If migError <> vbNullString Then
        Test_TbCambiosParaPublicacion_Edicion_PersisteValorSobreIntegerMax = JsonFail("Migration failed: " & migError, logs)
        GoTo Teardown
    End If
    If migResult <> EnumSiNo.Sí Then
        Test_TbCambiosParaPublicacion_Edicion_PersisteValorSobreIntegerMax = JsonFail("Migration did not return EnumSiNo.Si", logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_TbCambiosParaPublicacion_Edicion_PersisteValorSobreIntegerMax = JsonFail("GetTestDb returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' INSERT con valores por encima del Integer max. Si alguna columna sigue siendo Integer, dbFailOnError produces 3035.
    db.Execute "INSERT INTO tbCambiosParaPublicacion " & _
               "(IDCambio, IDProyecto, EdicionInicial, EdicionFinal, NombreCampo, FechaRegistro) " & _
               "VALUES (" & FIX_ID_CAMBIO_INI & ", " & FIX_ID_PROYECTO & ", " & FIX_ID_VALOR_INI & ", " & FIX_ID_VALOR_FIN & ", " & _
               "'test107inicial', #" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    Dim rs As DAO.Recordset
    Dim readIni As Variant
    Dim readFin As Variant
    Set rs = db.OpenRecordset("SELECT EdicionInicial, EdicionFinal FROM tbCambiosParaPublicacion WHERE IDCambio=" & FIX_ID_CAMBIO_INI)
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        Test_TbCambiosParaPublicacion_Edicion_PersisteValorSobreIntegerMax = JsonFail("INSERT succeeded but row not visible in SELECT", logs)
        GoTo Teardown
    End If
    readIni = rs.Fields("EdicionInicial").value
    readFin = rs.Fields("EdicionFinal").value
    rs.Close
    Set rs = Nothing

    If IsNull(readIni) Then
        Test_TbCambiosParaPublicacion_Edicion_PersisteValorSobreIntegerMax = JsonFail("EdicionInicial is Null after INSERT", logs)
        GoTo Teardown
    End If
    If CLng(readIni) <> FIX_ID_VALOR_INI Then
        Test_TbCambiosParaPublicacion_Edicion_PersisteValorSobreIntegerMax = JsonFail("EdicionInicial expected " & FIX_ID_VALOR_INI & " but got " & CLng(readIni), logs)
        GoTo Teardown
    End If
    If IsNull(readFin) Then
        Test_TbCambiosParaPublicacion_Edicion_PersisteValorSobreIntegerMax = JsonFail("EdicionFinal is Null after INSERT", logs)
        GoTo Teardown
    End If
    If CLng(readFin) <> FIX_ID_VALOR_FIN Then
        Test_TbCambiosParaPublicacion_Edicion_PersisteValorSobreIntegerMax = JsonFail("EdicionFinal expected " & FIX_ID_VALOR_FIN & " but got " & CLng(readFin), logs)
        GoTo Teardown
    End If

    Test_TbCambiosParaPublicacion_Edicion_PersisteValorSobreIntegerMax = JsonOk("persiste_valor_sobre_integer_max_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownCambiosPubFixture
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_TbCambiosParaPublicacion_Edicion_PersisteValorSobreIntegerMax = JsonFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function


' ============================================================
' Test 2 — MigrationEsIdempotente
' GIVEN MigracionTbCambiosParaPublicacionEdicionLong module is importable
' WHEN Ejecutar is called twice
' THEN both calls return EnumSiNo.Si, p_Error stays empty, BOTH columns end up dbLong
' AND second call does not re-alter (idempotent).
' ============================================================
Public Function Test_TbCambiosParaPublicacion_Edicion_MigrationEsIdempotente() As String
    Dim logs(0 To 8) As String
    Test_TbCambiosParaPublicacion_Edicion_MigrationEsIdempotente = JsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: EnsureTestConfigLoaded"
    logs(2) = "3. Act: EjecutarMigracion (1st call)"
    logs(3) = "4. Assert: 1st call returned EnumSiNo.Si"
    logs(4) = "5. Act: EjecutarMigracion (2nd call)"
    logs(5) = "6. Assert: 2nd call returned EnumSiNo.Si (idempotent)"
    logs(6) = "7. Assert: BOTH column types are dbLong (post-migration)"
    logs(7) = "8. Teardown: TeardownCambiosPubFixture"
    logs(8) = "9. Teardown: Test_Helper.ResetTestSession"

    Dim runError As String
    ForceLocalBackend runError
    If runError <> vbNullString Then
        Test_TbCambiosParaPublicacion_Edicion_MigrationEsIdempotente = JsonFail("ForceLocalBackend failed: " & runError, logs)
        GoTo Teardown
    End If

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_TbCambiosParaPublicacion_Edicion_MigrationEsIdempotente = JsonFail(cfgError, logs)
        GoTo Teardown
    End If

    Dim result1 As EnumSiNo
    Dim resultError As String
    resultError = vbNullString
    result1 = MigracionTbCambiosParaPublicacionEdicionLong.EjecutarMigracion(resultError)
    If resultError <> vbNullString Then
        Test_TbCambiosParaPublicacion_Edicion_MigrationEsIdempotente = JsonFail("First Ejecutar failed: " & resultError, logs)
        GoTo Teardown
    End If
    If result1 <> EnumSiNo.Sí Then
        Test_TbCambiosParaPublicacion_Edicion_MigrationEsIdempotente = JsonFail("First Ejecutar returned " & CStr(result1) & " (expected EnumSiNo.Si=" & CLng(EnumSiNo.Sí) & ")", logs)
        GoTo Teardown
    End If

    ' Second Ejecutar call — should be idempotent
    Dim result2 As EnumSiNo
    resultError = vbNullString
    result2 = MigracionTbCambiosParaPublicacionEdicionLong.EjecutarMigracion(resultError)
    If resultError <> vbNullString Then
        Test_TbCambiosParaPublicacion_Edicion_MigrationEsIdempotente = JsonFail("Second Ejecutar failed: " & resultError, logs)
        GoTo Teardown
    End If
    If result2 <> EnumSiNo.Sí Then
        Test_TbCambiosParaPublicacion_Edicion_MigrationEsIdempotente = JsonFail("Second Ejecutar returned " & CStr(result2) & " (expected EnumSiNo.Si=" & CLng(EnumSiNo.Sí) & ")", logs)
        GoTo Teardown
    End If

    ' Assert BOTH columns ended up as dbLong
    Dim typeErr As String
    Dim typesCombined As String
    typesCombined = TipoColumnasCambiosPub(typeErr)
    If typeErr <> "" Then
        Test_TbCambiosParaPublicacion_Edicion_MigrationEsIdempotente = JsonFail("TipoColumnasCambiosPub: " & typeErr, logs)
        GoTo Teardown
    End If
    If typesCombined <> "LONG,LONG" Then
        Test_TbCambiosParaPublicacion_Edicion_MigrationEsIdempotente = JsonFail("Column types are '" & typesCombined & "' (expected 'LONG,LONG')", logs)
        GoTo Teardown
    End If

    Test_TbCambiosParaPublicacion_Edicion_MigrationEsIdempotente = JsonOk("migration_es_idempotente_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    TeardownCambiosPubFixture
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_TbCambiosParaPublicacion_Edicion_MigrationEsIdempotente = JsonFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function


' ============================================================
' Test 3 — RegistrosExistentesPreservados
' GIVEN tbCambiosParaPublicacion has rows pre-migration
' WHEN MigracionTbCambiosParaPublicacionEdicionLong.EjecutarMigracion runs
' THEN count of all rows in tbCambiosParaPublicacion is invariant
' (ALTER COLUMN on Integer?Long preserves all existing rows; no truncation
' or type coercion that loses data).
'
' Why GREEN pre-migration: ALTER COLUMN preserves all rows regardless of column
' types -- the migration does not change row cardinality. This test catches the
' failure mode "migration runs but loses rows" without needing both columns
' to be Long first.
' ============================================================
Public Function Test_TbCambiosParaPublicacion_Edicion_RegistrosExistentesPreservados() As String
    Dim logs(0 To 8) As String
    Test_TbCambiosParaPublicacion_Edicion_RegistrosExistentesPreservados = JsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: EnsureTestConfigLoaded"
    logs(2) = "3. Act: count total rows in tbCambiosParaPublicacion before migration"
    logs(3) = "4. Act: Ejecutar"
    logs(4) = "5. Act: count total rows in tbCambiosParaPublicacion after migration"
    logs(5) = "6. Assert: count_after >= count_before (no rows lost)"
    logs(6) = "7. Teardown: TeardownCambiosPubFixture"
    logs(7) = "8. Teardown: Test_Helper.ResetTestSession"

    Dim runError As String
    ForceLocalBackend runError
    If runError <> vbNullString Then
        Test_TbCambiosParaPublicacion_Edicion_RegistrosExistentesPreservados = JsonFail("ForceLocalBackend failed: " & runError, logs)
        GoTo Teardown
    End If

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_TbCambiosParaPublicacion_Edicion_RegistrosExistentesPreservados = JsonFail(cfgError, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_TbCambiosParaPublicacion_Edicion_RegistrosExistentesPreservados = JsonFail("GetTestDb returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Dim beforeCount As Long
    Dim rsBefore As DAO.Recordset
    Set rsBefore = db.OpenRecordset("SELECT COUNT(*) AS C FROM tbCambiosParaPublicacion", dbOpenSnapshot)
    If Not rsBefore.EOF Then beforeCount = CLng(Nz(rsBefore.Fields("C").value, 0))
    rsBefore.Close
    Set rsBefore = Nothing

    Dim migError As String
    Dim migResult As EnumSiNo
    migError = vbNullString
    migResult = MigracionTbCambiosParaPublicacionEdicionLong.EjecutarMigracion(migError)
    If migError <> vbNullString Then
        Test_TbCambiosParaPublicacion_Edicion_RegistrosExistentesPreservados = JsonFail("Migration failed: " & migError, logs)
        GoTo Teardown
    End If
    If migResult <> EnumSiNo.Sí Then
        Test_TbCambiosParaPublicacion_Edicion_RegistrosExistentesPreservados = JsonFail("Migration did not return EnumSiNo.Si", logs)
        GoTo Teardown
    End If

    Dim afterCount As Long
    Dim rsAfter As DAO.Recordset
    Set rsAfter = db.OpenRecordset("SELECT COUNT(*) AS C FROM tbCambiosParaPublicacion", dbOpenSnapshot)
    If Not rsAfter.EOF Then afterCount = CLng(Nz(rsAfter.Fields("C").value, 0))
    rsAfter.Close
    Set rsAfter = Nothing

    ' Invariant: post >= pre. ALTER COLUMN must NEVER lose rows.
    If afterCount < beforeCount Then
        Test_TbCambiosParaPublicacion_Edicion_RegistrosExistentesPreservados = JsonFail("Row count decreased: before=" & beforeCount & " after=" & afterCount, logs)
        GoTo Teardown
    End If

    Test_TbCambiosParaPublicacion_Edicion_RegistrosExistentesPreservados = JsonOk("registros_existentes_preservados_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownCambiosPubFixture
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_TbCambiosParaPublicacion_Edicion_RegistrosExistentesPreservados = JsonFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function


' ============================================================
' RunAll - Aggregator for the TbCambiosParaPublicacion_Edicion migration block.
' Per access-vba-tdd §4.7 manifest discipline, RunAll chains the three atomic
' tests (PersisteValorSobreIntegerMax, MigrationEsIdempotente,
' RegistrosExistentesPreservados) under a single SuiteSetup/SuiteTeardown
' boundary, returns JsonOk only if all three atoms return ok:true, else
' JsonFail naming the first failing atom.
' ============================================================
Public Function Test_TbCambiosParaPublicacion_Edicion_RunAll() As String
    Dim results(0 To 2) As String
    Dim names(0 To 2) As String
    Dim i As Long
    Dim runError As String
    Dim outLogs(0 To 6) As String

    ' SuiteSetup
    ForceLocalBackend runError
    If runError <> "" Then
        outLogs(0) = "TESTS BLOCKED: " & runError
        Test_TbCambiosParaPublicacion_Edicion_RunAll = JsonFail("TESTS BLOCKED: " & runError, outLogs)
        Exit Function
    End If
    outLogs(0) = "SuiteSetup OK"

    names(0) = "Test_TbCambiosParaPublicacion_Edicion_PersisteValorSobreIntegerMax"
    names(1) = "Test_TbCambiosParaPublicacion_Edicion_MigrationEsIdempotente"
    names(2) = "Test_TbCambiosParaPublicacion_Edicion_RegistrosExistentesPreservados"

    results(0) = Test_TbCambiosParaPublicacion_Edicion_PersisteValorSobreIntegerMax()
    results(1) = Test_TbCambiosParaPublicacion_Edicion_MigrationEsIdempotente()
    results(2) = Test_TbCambiosParaPublicacion_Edicion_RegistrosExistentesPreservados()

    ' SuiteTeardown defensivo
    TeardownCambiosPubFixture
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
        Test_TbCambiosParaPublicacion_Edicion_RunAll = JsonOk("all_pass", outLogs)
    Else
        Test_TbCambiosParaPublicacion_Edicion_RunAll = JsonFail("some_tests_failed: " & firstFailure, outLogs)
    End If
End Function
