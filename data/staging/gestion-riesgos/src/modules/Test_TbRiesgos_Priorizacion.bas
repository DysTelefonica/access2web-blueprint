Attribute VB_Name = "Test_TbRiesgos_Priorizacion"
Option Compare Database
Option Explicit

' ============================================================
' Test_TbRiesgos_Priorizacion -- Tests para el fix de issue #106
'
' Cobertura (3 tests atomicos, REQ-MIG-106-01..03):
'   Test 1 -- PersisteValorSobreIntegerMax (controlled cold temp backend starts
'           as Integer, migrates to Long, then DAO accepts Priorizacion=50000)
'   Test 2 -- MigrationEsIdempotente (MigracionRiesgosPriorizacionLong.EjecutarMigracion
'           called twice returns EnumSiNo.Si both times; RED pre-migration
'           because helper does not exist yet -- late-bound via Application.Run)
'   Test 3 -- RegistrosExistentesPreservados (count of rows with Priorizacion IS NOT NULL
'           unchanged after migration; GREEN from start because migration is
'           cardinality-invariant regardless of column type)
' ============================================================

' --- Module-level declarations (AGENTS.md rule 3) ---
Private Const FIX_ID_BASE As Long = 901060
Private Const FIX_ID_RIESGO As Long = 901061
Private Const FIX_ID_EDICION As Long = 901062
Private Const FIX_ID_PROYECTO As Long = 901063
Private Const FIX_ID_EXPEDIENTE As Long = 901064
Private Const FIX_PRIORIZACION_VALOR As Long = 50000

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

' --- GetTestDb (delegacion a Test_Fixtures) ---
Private Function GetTestDb(ByRef p_Error As String) As DAO.Database
    Set GetTestDb = Test_Fixtures.GetTestDb(p_Error)
End Function

' --- Helper: read current Priorizacion column type via DAO ---
' Returns the DAO dbType constant (dbInteger=3, dbLong=4). Empty p_Error on success.
Private Function TipoColumnaPriorizacion(ByRef p_Error As String) As Long
    On Error GoTo EH
    Dim db As DAO.Database
    Dim dbErr As String
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        p_Error = "TipoColumnaPriorizacion: GetTestDb returned Nothing: " & dbErr
        Exit Function
    End If

    Set tdf = db.TableDefs("TbRiesgos")
    Set fld = tdf.Fields("Priorizacion")
    TipoColumnaPriorizacion = fld.Type

    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

EH:
    p_Error = "TipoColumnaPriorizacion: " & Err.Description
End Function

Private Function TipoColumnaPriorizacionInDb( _
    ByVal p_db As DAO.Database, _
    ByRef p_Error As String _
) As Long
    On Error GoTo EH
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    Set tdf = p_db.TableDefs("TbRiesgos")
    Set fld = tdf.Fields("Priorizacion")
    TipoColumnaPriorizacionInDb = fld.Type

    Set fld = Nothing
    Set tdf = Nothing
    Exit Function

EH:
    p_Error = "TipoColumnaPriorizacionInDb: " & Err.Description
End Function

Private Function CreateColdPriorizacionBackend( _
    ByRef p_db As DAO.Database, _
    ByRef p_Path As String, _
    ByRef p_Error As String _
) As Boolean
    On Error GoTo EH
    Dim fso As Object
    Dim tempFolder As String
    Dim tdf As DAO.TableDef
    Dim fldId As DAO.Field
    Dim fldPriorizacion As DAO.Field

    Randomize
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempFolder = Environ$("TEMP")
    If tempFolder = "" Then tempFolder = fso.GetSpecialFolder(2).Path
    p_Path = tempFolder & "\GestionRiesgos_Issue106_" & _
             Format$(Now, "yyyymmddhhnnss") & "_" & _
             CStr(Int(Rnd() * 100000)) & ".accdb"

    Set p_db = DBEngine.Workspaces(0).CreateDatabase(p_Path, dbLangGeneral, dbVersion120)

    Set tdf = p_db.CreateTableDef("TbRiesgos")
    Set fldId = tdf.CreateField("IDRiesgo", dbLong)
    Set fldPriorizacion = tdf.CreateField("Priorizacion", dbInteger)
    tdf.Fields.Append fldId
    tdf.Fields.Append fldPriorizacion
    p_db.TableDefs.Append tdf
    p_db.TableDefs.Refresh

    p_db.Execute "INSERT INTO TbRiesgos (IDRiesgo, Priorizacion) VALUES (" & FIX_ID_RIESGO & ", 100)", dbFailOnError

    CreateColdPriorizacionBackend = True
    Set fldPriorizacion = Nothing
    Set fldId = Nothing
    Set tdf = Nothing
    Set fso = Nothing
    Exit Function

EH:
    p_Error = "CreateColdPriorizacionBackend: " & Err.Description
    On Error Resume Next
    If Not p_db Is Nothing Then
        p_db.Close
        Set p_db = Nothing
    End If
    If p_Path <> "" Then
        If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(p_Path) Then fso.DeleteFile p_Path, True
    End If
    Set fso = Nothing
    On Error GoTo 0
End Function

Private Sub CleanupTempBackend(ByRef p_db As DAO.Database, ByVal p_Path As String)
    On Error Resume Next
    Dim fso As Object
    If Not p_db Is Nothing Then
        p_db.Close
        Set p_db = Nothing
    End If
    If p_Path <> "" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(p_Path) Then fso.DeleteFile p_Path, True
        Set fso = Nothing
    End If
    On Error GoTo 0
End Sub

' --- Fixture helpers (FK-ordered) ---
Private Sub SeedRiesgosFixture()
    On Error GoTo EH_Seed
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then Err.Raise 1001, "SeedRiesgosFixture", "GetTestDb returned Nothing: " & dbErr

    ' Idempotent cleanup in reverse FK order (only deterministic test markers)
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & FIX_ID_RIESGO, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_ID_PROYECTO, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_ID_EXPEDIENTE, dbFailOnError
    On Error GoTo 0

    ' 1. TbExpedientes (padre)
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
               "VALUES (" & FIX_ID_EXPEDIENTE & ", 'TEST106', 'Fixture issue 106', 'Test', 1)", dbFailOnError

    ' 2. TbProyectos (hijo de Expediente)
    db.Execute "INSERT INTO TbProyectos " & _
               "(IDProyecto, IDExpediente, Proyecto, ParaInformeAvisos, NombreUsuarioCalidad) " & _
               "VALUES (" & FIX_ID_PROYECTO & ", " & FIX_ID_EXPEDIENTE & ", 'TESTPROJ106', 'Si', 'test_calidad_106')", dbFailOnError

    ' 3. TbProyectosEdiciones (hijo de Proyecto)
    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDEdicion, IDProyecto, Edicion, Elaborado, FechaPreparadaParaPublicar) " & _
               "VALUES (" & FIX_ID_EDICION & ", " & FIX_ID_PROYECTO & ", 1, " & _
               "'test_calidad_106', #" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    ' Post-seed cardinality assertion -- guard against verde-por-suerte.
    Dim seedRs As DAO.Recordset
    Dim seedCount As Long
    Set seedRs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_ID_EDICION)
    If Not seedRs.EOF Then seedCount = CLng(Nz(seedRs.Fields("C").Value, 0))
    seedRs.Close
    Set seedRs = Nothing
    If seedCount <> 1 Then
        Err.Raise 1002, "SeedRiesgosFixture", "Post-seed assertion failed: expected 1 row in TbProyectosEdiciones, got " & seedCount
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
    Err.Raise eNuma, "SeedRiesgosFixture", "Seed failed: " & eNuma & " - " & eDesc
End Sub

Private Sub TeardownRiesgosFixture()
    On Error Resume Next
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then Exit Sub

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & FIX_ID_RIESGO, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_ID_PROYECTO, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_ID_EXPEDIENTE, dbFailOnError

    Set db = Nothing
    On Error GoTo 0
End Sub

' ============================================================
' Test 1 -- PersisteValorSobreIntegerMax (REQ-MIG-106-01)
' ============================================================
Public Function Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax() As String
    Dim logs(0 To 10) As String
    Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax = JsonFail("test did not complete", logs)
    On Error GoTo HandleError

    Dim db As DAO.Database
    Dim tempPath As String
    Dim tempErr As String

    logs(0) = "1. Arrange: create isolated temp backend with Priorizacion INTEGER"
    logs(1) = "2. Assert: cold schema starts as dbInteger"
    logs(2) = "3. Act: EjecutarMigracionEnBase against the controlled temp backend"
    logs(3) = "4. Assert: schema is dbLong after migration"
    logs(4) = "5. Act: INSERT TbRiesgos (Priorizacion=" & FIX_PRIORIZACION_VALOR & ")"
    logs(5) = "6. Assert: read Priorizacion back = " & FIX_PRIORIZACION_VALOR
    logs(6) = "7. Assert: fixture cardinality increased by one"
    logs(7) = "8. Teardown: delete isolated temp backend"

    If Not CreateColdPriorizacionBackend(db, tempPath, tempErr) Then
        Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax = JsonFail(tempErr, logs)
        GoTo Teardown
    End If

    Dim colTypeErr As String
    Dim colType As Long
    colType = TipoColumnaPriorizacionInDb(db, colTypeErr)
    If colTypeErr <> "" Then
        Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax = JsonFail(colTypeErr, logs)
        GoTo Teardown
    End If
    If colType <> dbInteger Then
        Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax = JsonFail("Cold schema expected dbInteger=" & dbInteger & " but got " & colType, logs)
        GoTo Teardown
    End If

    Dim migError As String
    Dim migResult As EnumSiNo
    migResult = MigracionRiesgosPriorizacionLong.EjecutarMigracionEnBase(db, migError)
    If migError <> "" Then
        Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax = JsonFail("Migration failed: " & migError, logs)
        GoTo Teardown
    End If
    If migResult <> EnumSiNo.Sí Then
        Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax = JsonFail("Migration returned EnumSiNo.No", logs)
        GoTo Teardown
    End If

    colTypeErr = ""
    colType = TipoColumnaPriorizacionInDb(db, colTypeErr)
    If colTypeErr <> "" Then
        Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax = JsonFail(colTypeErr, logs)
        GoTo Teardown
    End If
    If colType <> dbLong Then
        Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax = JsonFail("Post-migration schema expected dbLong=" & dbLong & " but got " & colType, logs)
        GoTo Teardown
    End If

    Dim countBefore As Long
    Dim countAfter As Long
    db.Execute "INSERT INTO TbRiesgos " & _
               "(IDRiesgo, Priorizacion) " & _
               "VALUES (" & (FIX_ID_RIESGO + 1) & ", " & FIX_PRIORIZACION_VALOR & ")", dbFailOnError

    Dim rs As DAO.Recordset
    Dim readValue As Variant
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbRiesgos WHERE IDRiesgo IN (" & FIX_ID_RIESGO & ", " & (FIX_ID_RIESGO + 1) & ")")
    If Not rs.EOF Then countAfter = CLng(Nz(rs.Fields("C").Value, 0))
    rs.Close
    Set rs = Nothing

    countBefore = 1
    If countAfter <> countBefore + 1 Then
        Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax = JsonFail("Cardinality expected " & (countBefore + 1) & " fixture rows but got " & countAfter, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset("SELECT TOP 1 Priorizacion FROM TbRiesgos WHERE IDRiesgo=" & (FIX_ID_RIESGO + 1))
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax = JsonFail("INSERT succeeded but row not visible in SELECT", logs)
        GoTo Teardown
    End If
    readValue = rs.Fields("Priorizacion").Value
    rs.Close
    Set rs = Nothing

    If IsNull(readValue) Then
        Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax = JsonFail("Priorizacion is Null after INSERT", logs)
        GoTo Teardown
    End If
    If CLng(readValue) <> FIX_PRIORIZACION_VALOR Then
        Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax = JsonFail("Priorizacion expected " & FIX_PRIORIZACION_VALOR & " but got " & CLng(readValue), logs)
        GoTo Teardown
    End If

    Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax = JsonOk("persiste_valor_sobre_integer_max_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    CleanupTempBackend db, tempPath
    On Error GoTo 0
    Exit Function

HandleError:
    Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax = JsonFail("unexpected: " & Err.Description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' Test 2 -- MigrationEsIdempotente (REQ-MIG-106-01 + REQ-MIG-106-03)
' ============================================================
Public Function Test_TbRiesgos_Priorizacion_MigrationEsIdempotente() As String
    Dim logs(0 To 8) As String
    Test_TbRiesgos_Priorizacion_MigrationEsIdempotente = JsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Act: Ejecutar (1st direct public call)"
    logs(2) = "3. Assert: 1st call returned EnumSiNo.Si"
    logs(3) = "4. Act: Ejecutar (2nd call)"
    logs(4) = "5. Assert: 2nd call returned EnumSiNo.Si (idempotent)"
    logs(5) = "6. Assert: column type is dbLong (post-migration)"
    logs(6) = "7. Teardown: TeardownRiesgosFixture"
    logs(7) = "8. Teardown: Test_Helper.ResetTestSession"

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_TbRiesgos_Priorizacion_MigrationEsIdempotente = JsonFail(cfgError, logs)
        GoTo Teardown
    End If

    Dim result1 As EnumSiNo
    Dim migError As String
    result1 = MigracionRiesgosPriorizacionLong.EjecutarMigracion(migError)
    If migError <> "" Then
        Test_TbRiesgos_Priorizacion_MigrationEsIdempotente = JsonFail("First Ejecutar failed: " & migError, logs)
        GoTo Teardown
    End If
    If CLng(result1) <> CLng(EnumSiNo.Sí) Then
        Test_TbRiesgos_Priorizacion_MigrationEsIdempotente = JsonFail("First Ejecutar returned " & CStr(result1) & " (expected EnumSiNo.Si=" & CLng(EnumSiNo.Sí) & ")", logs)
        GoTo Teardown
    End If

    Dim result2 As EnumSiNo
    migError = ""
    result2 = MigracionRiesgosPriorizacionLong.EjecutarMigracion(migError)
    If migError <> "" Then
        Test_TbRiesgos_Priorizacion_MigrationEsIdempotente = JsonFail("Second Ejecutar failed: " & migError, logs)
        GoTo Teardown
    End If
    If CLng(result2) <> CLng(EnumSiNo.Sí) Then
        Test_TbRiesgos_Priorizacion_MigrationEsIdempotente = JsonFail("Second Ejecutar returned " & CStr(result2) & " (expected EnumSiNo.Si=" & CLng(EnumSiNo.Sí) & ")", logs)
        GoTo Teardown
    End If

    Dim colTypeErr As String
    Dim colType As Long
    colType = TipoColumnaPriorizacion(colTypeErr)
    If colTypeErr <> "" Then
        Test_TbRiesgos_Priorizacion_MigrationEsIdempotente = JsonFail("TipoColumnaPriorizacion: " & colTypeErr, logs)
        GoTo Teardown
    End If
    If colType <> dbLong Then
        Test_TbRiesgos_Priorizacion_MigrationEsIdempotente = JsonFail("Column type is " & colType & " (expected dbLong=" & dbLong & ")", logs)
        GoTo Teardown
    End If

    Test_TbRiesgos_Priorizacion_MigrationEsIdempotente = JsonOk("migration_es_idempotente_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    TeardownRiesgosFixture
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_TbRiesgos_Priorizacion_MigrationEsIdempotente = JsonFail("unexpected: " & Err.Description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' Test 3 -- RegistrosExistentesPreservados (REQ-MIG-106-03)
' ============================================================
Public Function Test_TbRiesgos_Priorizacion_RegistrosExistentesPreservados() As String
    Dim logs(0 To 8) As String
    Test_TbRiesgos_Priorizacion_RegistrosExistentesPreservados = JsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Act: count(TbRiesgos WHERE Priorizacion IS NOT NULL) before migration"
    logs(2) = "3. Act: Ejecutar"
    logs(3) = "4. Act: count after migration"
    logs(4) = "5. Assert: count_after = count_before (schema-only migration preserves exact cardinality)"
    logs(5) = "6. Teardown: TeardownRiesgosFixture"
    logs(6) = "7. Teardown: Test_Helper.ResetTestSession"

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_TbRiesgos_Priorizacion_RegistrosExistentesPreservados = JsonFail(cfgError, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_TbRiesgos_Priorizacion_RegistrosExistentesPreservados = JsonFail("GetTestDb returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Dim countBefore As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbRiesgos WHERE Priorizacion IS NOT NULL")
    If Not rs.EOF Then countBefore = CLng(Nz(rs.Fields("C").Value, 0))
    rs.Close
    Set rs = Nothing
    logs(1) = "2. count_before (global, not just fixture): " & countBefore

    Dim migResult As EnumSiNo
    Dim migError As String
    migResult = MigracionRiesgosPriorizacionLong.EjecutarMigracion(migError)
    If migError <> "" Then
        Test_TbRiesgos_Priorizacion_RegistrosExistentesPreservados = JsonFail("Ejecutar failed: " & migError, logs)
        GoTo Teardown
    End If
    If CLng(migResult) <> CLng(EnumSiNo.Sí) Then
        Test_TbRiesgos_Priorizacion_RegistrosExistentesPreservados = JsonFail("Ejecutar returned " & CStr(migResult) & " (expected EnumSiNo.Si=" & CLng(EnumSiNo.Sí) & ")", logs)
        GoTo Teardown
    End If

    Dim countAfter As Long
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbRiesgos WHERE Priorizacion IS NOT NULL")
    If Not rs.EOF Then countAfter = CLng(Nz(rs.Fields("C").Value, 0))
    rs.Close
    Set rs = Nothing
    logs(3) = "4. count_after (global, not just fixture): " & countAfter

    If countAfter <> countBefore Then
        Test_TbRiesgos_Priorizacion_RegistrosExistentesPreservados = JsonFail("Cardinality changed during schema-only migration: count_before=" & countBefore & " count_after=" & countAfter, logs)
        GoTo Teardown
    End If

    Test_TbRiesgos_Priorizacion_RegistrosExistentesPreservados = JsonOk("registros_existentes_preservados_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownRiesgosFixture
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_TbRiesgos_Priorizacion_RegistrosExistentesPreservados = JsonFail("unexpected: " & Err.Description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' RunAll -- Aggregator. Per access-vba-tdd manifest discipline:
' RunAll goes in smoke or direct execution, not in the atomic manifest.
' ============================================================
Public Function Test_TbRiesgos_Priorizacion_RunAll() As String
    Dim results(0 To 2) As String
    Dim names(0 To 2) As String
    Dim i As Long
    Dim runError As String
    Dim outLogs(0 To 6) As String

    ForceLocalBackend runError
    If runError <> "" Then
        outLogs(0) = "TESTS BLOCKED: " & runError
        Test_TbRiesgos_Priorizacion_RunAll = JsonFail("TESTS BLOCKED: " & runError, outLogs)
        Exit Function
    End If
    outLogs(0) = "SuiteSetup OK"

    names(0) = "Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax"
    names(1) = "Test_TbRiesgos_Priorizacion_MigrationEsIdempotente"
    names(2) = "Test_TbRiesgos_Priorizacion_RegistrosExistentesPreservados"

    results(0) = Test_TbRiesgos_Priorizacion_PersisteValorSobreIntegerMax()
    results(1) = Test_TbRiesgos_Priorizacion_MigrationEsIdempotente()
    results(2) = Test_TbRiesgos_Priorizacion_RegistrosExistentesPreservados()

    TeardownRiesgosFixture
    outLogs(1) = "SuiteTeardown OK"

    Test_Helper.ResetTestSession
    outLogs(2) = "ResetTestSession OK"

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
        Test_TbRiesgos_Priorizacion_RunAll = JsonOk("all_pass", outLogs)
    Else
        Test_TbRiesgos_Priorizacion_RunAll = JsonFail("some_tests_failed: " & firstFailure, outLogs)
    End If
End Function
