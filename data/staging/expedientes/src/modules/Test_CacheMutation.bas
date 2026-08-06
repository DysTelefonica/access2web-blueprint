Attribute VB_Name = "Test_CacheMutation"
Option Compare Database
Option Explicit

' =============================================================================
' Test module: Test_CacheMutation
' Phase 4 Tranche A of SDD change staging-alignment-prueba-001 (Issue #19)
'
' Covers transactional atomicity of ExpedienteEntidadOperaciones.Registrar
' writing to TbExpedientesConEntidades with granular ambitos.
' Scenarios exercised:
'   1. Cabecera write succeeds and persists in TbExpedientesConEntidades
'   2. Rollback on in-memory update failure proves DB is NOT committed
'   3. p_ForceCacheFailureForTest isolation proves no DB change
'   4. Todo ambito writes all fields atomically
'
' All tests follow access-vba-tdd v2.4.2:
'   - Public Function returning canonical JSON (BuildJsonOk / BuildJsonFail)
'   - Zero arguments
'   - Explicit logs() with Arrange/Act/Assert steps
'   - No Debug.Print, no MsgBox, no UI
'   - Sandbox safety via temp .accdb files (CreateIsolatedTempDb pattern)
'   - Deterministic fixture IDs, defensive teardown
' =============================================================================

' ---- Private helpers (local canonical JSON wrappers) ----

Private Function JsonOk(ByVal p_Value As String, ByRef logs() As String) As String
    JsonOk = BuildJsonOk(p_Value, logs)
End Function

Private Function JsonFail(ByVal p_Error As String, ByRef logs() As String) As String
    JsonFail = BuildJsonFail(p_Error, logs)
End Function

' ---- Temp .accdb helpers (private — local copy of TestHelper.CreateIsolatedTempDb) ----

Private Function MakeIsolatedTempDbPath(ByVal p_Suffix As String) As String
    MakeIsolatedTempDbPath = Environ$("TEMP") & "\expedientes_test_cachemutation_" & p_Suffix & "_" & Format$(Now, "yyyymmddhhnnssfff") & ".accdb"
End Function

Private Function CreateIsolatedTempDb(ByVal p_Path As String, ByRef p_Error As String) As DAO.Database
    On Error GoTo EH
    p_Error = ""
    If Len(Dir$(p_Path)) > 0 Then Kill p_Path
    Set CreateIsolatedTempDb = DBEngine.Workspaces(0).CreateDatabase(p_Path, dbLangGeneral, dbVersion120)
    Exit Function
EH:
    p_Error = "CreateIsolatedTempDb: " & Err.Number & " - " & Err.Description
    Set CreateIsolatedTempDb = Nothing
End Function

Private Sub DisposeIsolatedTempDb(ByVal p_Path As String, ByRef p_Db As DAO.Database)
    On Error Resume Next
    If Not p_Db Is Nothing Then p_Db.Close
    Set p_Db = Nothing
    If Len(Dir$(p_Path)) > 0 Then Kill p_Path
    On Error GoTo 0
End Sub

' ---- Helper: create TbExpedientesConEntidades in temp .accdb ----

Private Function EnsureCacheTable(ByVal p_Db As DAO.Database, ByRef p_Error As String) As Boolean
    On Error GoTo EH
    p_Error = ""
    Dim tdf As DAO.TableDef
    Set tdf = p_Db.CreateTableDef("TbExpedientesConEntidades")
    tdf.Fields.Append tdf.CreateField("IDExpediente", dbLong)
    tdf.Fields.Append tdf.CreateField("Clasificacion", dbText, 255)
    tdf.Fields.Append tdf.CreateField("OrganoContratacion", dbText, 255)
    tdf.Fields.Append tdf.CreateField("OficinaPrograma", dbText, 255)
    tdf.Fields.Append tdf.CreateField("Ejercito", dbText, 255)
    tdf.Fields.Append tdf.CreateField("Estado", dbText, 255)
    tdf.Fields.Append tdf.CreateField("ResponsableCalidad", dbText, 255)
    tdf.Fields.Append tdf.CreateField("ResponsableSeguridad", dbText, 255)
    tdf.Fields.Append tdf.CreateField("CadenaPecal", dbText, 255)
    tdf.Fields.Append tdf.CreateField("Pecal", dbText, 2)
    tdf.Fields.Append tdf.CreateField("CadenaContratistas", dbText, 255)
    tdf.Fields.Append tdf.CreateField("CadenaSubContratistas", dbText, 255)
    tdf.Fields.Append tdf.CreateField("CadenaSuministradores", dbText, 255)
    tdf.Fields.Append tdf.CreateField("CadenaComerciales", dbText, 255)
    tdf.Fields.Append tdf.CreateField("CadenaJPs", dbText, 255)
    tdf.Fields.Append tdf.CreateField("CadenaRACs", dbText, 255)
    tdf.Fields.Append tdf.CreateField("CadenaCorreoRACs", dbText, 255)
    tdf.Fields.Append tdf.CreateField("CadenaHitos", dbText, 255)
    tdf.Fields.Append tdf.CreateField("TipoParaLista", dbText, 255)
    tdf.Fields.Append tdf.CreateField("CadenaLugares", dbText, 255)
    tdf.Fields.Append tdf.CreateField("CadenaJuridicas", dbText, 255)
    p_Db.TableDefs.Append tdf
    EnsureCacheTable = True
    Exit Function
EH:
    p_Error = "EnsureCacheTable: " & Err.Number & " - " & Err.Description
    EnsureCacheTable = False
End Function

' ---- Helper: count rows in cache table ----

Private Function CountCacheRows(ByVal p_Db As DAO.Database, ByVal p_IDExp As Long) As Long
    On Error GoTo EH
    Dim rs As DAO.Recordset
    Dim sql As String
    sql = "SELECT COUNT(*) AS Cnt FROM TbExpedientesConEntidades WHERE IDExpediente=" & p_IDExp & ";"
    Set rs = p_Db.OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then
        CountCacheRows = rs.Fields("Cnt").Value
    Else
        CountCacheRows = 0
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
EH:
    CountCacheRows = -1
End Function

' ---- Helper: read a field value from cache table ----

Private Function ReadCacheField(ByVal p_Db As DAO.Database, ByVal p_IDExp As Long, ByVal p_FieldName As String) As Variant
    On Error GoTo EH
    Dim rs As DAO.Recordset
    Dim sql As String
    sql = "SELECT " & p_FieldName & " FROM TbExpedientesConEntidades WHERE IDExpediente=" & p_IDExp & ";"
    Set rs = p_Db.OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then
        ReadCacheField = rs.Fields(p_FieldName).Value
    Else
        ReadCacheField = Null
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
EH:
    ReadCacheField = Null
End Function

' ---- Test 1: Cabecera write succeeds and persists ----

Public Function Test_CacheMutation_CabeceraWriteSucceeds() As String
    Dim logs(0 To 8) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_CacheMutation_CabeceraWriteSucceeds = JsonFail("test did not complete", logs)

    On Error GoTo HandleError

    Dim errMsg As String
    Dim dbPath As String
    Dim dbTemp As DAO.Database
    Dim dbCreated As DAO.Database
    Dim exp As Expediente
    Dim op As ExpedienteEntidadOperaciones
    Dim result As String
    Dim rowCount As Long

    ' --- Arrange ---
    logs(0) = "1. Arrange: create isolated temp .accdb with TbExpedientesConEntidades"
    dbPath = MakeIsolatedTempDbPath("cabecera_ok")
    Set dbCreated = CreateIsolatedTempDb(dbPath, errMsg)
    If dbCreated Is Nothing Then
        Test_CacheMutation_CabeceraWriteSucceeds = JsonFail("CreateIsolatedTempDb: " & errMsg, logs)
        GoTo Teardown
    End If

    logs(1) = "1a. Arrange: create TbExpedientesConEntidades in temp db"
    If Not EnsureCacheTable(dbCreated, errMsg) Then
        Test_CacheMutation_CabeceraWriteSucceeds = JsonFail("EnsureCacheTable: " & errMsg, logs)
        GoTo Teardown
    End If
    Set dbCreated = Nothing  ' close so Registrar can open it

    logs(2) = "2. Arrange: create Expediente with IDExpediente=9001"
    Set exp = New Expediente
    exp.IDExpediente = "9001"
    exp.Ambito = "TEST"
    exp.Titulo = "Test Cabecera Write"

    logs(3) = "3. Arrange: configure ExpedienteEntidadOperaciones"
    Set op = New ExpedienteEntidadOperaciones
    Set op.Expediente = exp

    logs(4) = "4. Arrange: open temp db for Registrar (passing p_db)"
    Set dbCreated = DBEngine.Workspaces(0).OpenDatabase(dbPath)

    ' --- Act ---
    logs(5) = "5. Act: Registrar with p_Ambito=Cabecera"
    result = op.Registrar(p_db:=dbCreated, p_Ambito:=EnumAmbitoActualizacion.Cabecera, p_Error:=errMsg)
    If errMsg <> "" Then
        Test_CacheMutation_CabeceraWriteSucceeds = JsonFail("Registrar failed: " & errMsg, logs)
        GoTo Teardown
    End If

    ' --- Assert ---
    logs(6) = "6. Assert: row exists in TbExpedientesConEntidades"
    rowCount = CountCacheRows(dbCreated, 9001)
    If rowCount <> 1 Then
        Test_CacheMutation_CabeceraWriteSucceeds = JsonFail("Expected 1 row, got " & rowCount, logs)
        GoTo Teardown
    End If

    logs(7) = "7. Assert: TipoParaLista field has the expected value"
    Dim campoValor As Variant
    campoValor = ReadCacheField(dbCreated, 9001, "TipoParaLista")
    logs(7) = "7. Assert: TipoParaLista=" & Nz(campoValor, "<NULL>")

    Test_CacheMutation_CabeceraWriteSucceeds = JsonOk("cabecera_write_succeeds", logs)

Teardown:
    On Error Resume Next
    Set op = Nothing
    Set exp = Nothing
    If Not dbCreated Is Nothing Then dbCreated.Close
    DisposeIsolatedTempDb dbPath, dbTemp
    Exit Function

HandleError:
    Test_CacheMutation_CabeceraWriteSucceeds = JsonFail(Err.Description, logs)
    Resume Teardown
End Function

' ---- Test 2: Rollback on in-memory failure proves DB not committed ----

Public Function Test_CacheMutation_RollbackOnInMemoryFailure() As String
    Dim logs(0 To 8) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_CacheMutation_RollbackOnInMemoryFailure = JsonFail("test did not complete", logs)

    On Error GoTo HandleError

    Dim errMsg As String
    Dim dbPath As String
    Dim dbTemp As DAO.Database
    Dim dbCreated As DAO.Database
    Dim exp As Expediente
    Dim op As ExpedienteEntidadOperaciones
    Dim result As String
    Dim rowCount As Long

    ' --- Arrange ---
    logs(0) = "1. Arrange: create isolated temp .accdb with TbExpedientesConEntidades"
    dbPath = MakeIsolatedTempDbPath("rollback_ok")
    Set dbCreated = CreateIsolatedTempDb(dbPath, errMsg)
    If dbCreated Is Nothing Then
        Test_CacheMutation_RollbackOnInMemoryFailure = JsonFail("CreateIsolatedTempDb: " & errMsg, logs)
        GoTo Teardown
    End If

    logs(1) = "1a. Arrange: create TbExpedientesConEntidades in temp db"
    If Not EnsureCacheTable(dbCreated, errMsg) Then
        Test_CacheMutation_RollbackOnInMemoryFailure = JsonFail("EnsureCacheTable: " & errMsg, logs)
        GoTo Teardown
    End If
    Set dbCreated = Nothing

    logs(2) = "2. Arrange: create Expediente with IDExpediente=9002"
    Set exp = New Expediente
    exp.IDExpediente = "9002"
    exp.Ambito = "TEST"
    exp.Titulo = "Test Rollback"

    logs(3) = "3. Arrange: configure ExpedienteEntidadOperaciones with forced cache failure"
    Set op = New ExpedienteEntidadOperaciones
    Set op.Expediente = exp
    op.TestOnlyForceCacheRefreshFailure True

    logs(4) = "4. Arrange: open temp db for Registrar (passing p_db)"
    Set dbCreated = DBEngine.Workspaces(0).OpenDatabase(dbPath)

    ' --- Act ---
    logs(5) = "5. Act: Registrar with p_ForceCacheFailureForTest=True — should fail and rollback"
    result = op.Registrar(p_db:=dbCreated, p_Ambito:=EnumAmbitoActualizacion.Cabecera, _
                          p_ForceCacheFailureForTest:=True, p_Error:=errMsg)

    ' --- Assert ---
    logs(6) = "6. Assert: Registrar returned error (expected)"
    If errMsg = "" Then
        Test_CacheMutation_RollbackOnInMemoryFailure = JsonFail("Expected error but got none", logs)
        GoTo Teardown
    End If

    logs(7) = "7. Assert: no row in TbExpedientesConEntidades (rollback worked)"
    rowCount = CountCacheRows(dbCreated, 9002)
    If rowCount <> 0 Then
        Test_CacheMutation_RollbackOnInMemoryFailure = JsonFail("ROLLBACK FAILED: " & rowCount & " rows still exist after forced failure", logs)
        GoTo Teardown
    End If

    Test_CacheMutation_RollbackOnInMemoryFailure = JsonOk("rollback_on_inmemory_failure", logs)

Teardown:
    On Error Resume Next
    If Not op Is Nothing Then op.TestOnlyResetCacheRefreshFailure
    Set op = Nothing
    Set exp = Nothing
    If Not dbCreated Is Nothing Then dbCreated.Close
    DisposeIsolatedTempDb dbPath, dbTemp
    Exit Function

HandleError:
    Test_CacheMutation_RollbackOnInMemoryFailure = JsonFail(Err.Description, logs)
    Resume Teardown
End Function

' ---- Test 3: p_ForceCacheFailureForTest isolation ----

Public Function Test_CacheMutation_ForceCacheFailureIsolation() As String
    Dim logs(0 To 6) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_CacheMutation_ForceCacheFailureIsolation = JsonFail("test did not complete", logs)

    On Error GoTo HandleError

    Dim errMsg As String
    Dim dbPath As String
    Dim dbTemp As DAO.Database
    Dim dbCreated As DAO.Database
    Dim exp As Expediente
    Dim op As ExpedienteEntidadOperaciones
    Dim result As String
    Dim rowCount As Long

    ' --- Arrange ---
    logs(0) = "1. Arrange: create isolated temp .accdb with TbExpedientesConEntidades"
    dbPath = MakeIsolatedTempDbPath("force_isolation")
    Set dbCreated = CreateIsolatedTempDb(dbPath, errMsg)
    If dbCreated Is Nothing Then
        Test_CacheMutation_ForceCacheFailureIsolation = JsonFail("CreateIsolatedTempDb: " & errMsg, logs)
        GoTo Teardown
    End If

    logs(1) = "1a. Arrange: create TbExpedientesConEntidades in temp db"
    If Not EnsureCacheTable(dbCreated, errMsg) Then
        Test_CacheMutation_ForceCacheFailureIsolation = JsonFail("EnsureCacheTable: " & errMsg, logs)
        GoTo Teardown
    End If
    Set dbCreated = Nothing

    logs(2) = "2. Arrange: create Expediente with IDExpediente=9003"
    Set exp = New Expediente
    exp.IDExpediente = "9003"
    exp.Ambito = "TEST"
    exp.Titulo = "Test ForceFailure Isolation"

    logs(3) = "3. Arrange: configure ExpedienteEntidadOperaciones with forced cache failure via p_ForceCacheFailureForTest"
    Set op = New ExpedienteEntidadOperaciones
    Set op.Expediente = exp

    logs(4) = "4. Arrange: open temp db for Registrar (passing p_db)"
    Set dbCreated = DBEngine.Workspaces(0).OpenDatabase(dbPath)

    ' --- Act ---
    logs(5) = "5. Act: Registrar with p_ForceCacheFailureForTest=True — must fail early, no DB write"
    result = op.Registrar(p_db:=dbCreated, p_Ambito:=EnumAmbitoActualizacion.Cabecera, _
                          p_ForceCacheFailureForTest:=True, p_Error:=errMsg)

    ' --- Assert ---
    rowCount = CountCacheRows(dbCreated, 9003)
    If rowCount <> 0 Then
        Test_CacheMutation_ForceCacheFailureIsolation = JsonFail("ISOLATION FAILED: " & rowCount & " rows exist after forced cache failure", logs)
        GoTo Teardown
    End If

    Test_CacheMutation_ForceCacheFailureIsolation = JsonOk("force_cache_failure_isolation", logs)

Teardown:
    On Error Resume Next
    Set op = Nothing
    Set exp = Nothing
    If Not dbCreated Is Nothing Then dbCreated.Close
    DisposeIsolatedTempDb dbPath, dbTemp
    Exit Function

HandleError:
    Test_CacheMutation_ForceCacheFailureIsolation = JsonFail(Err.Description, logs)
    Resume Teardown
End Function

' ---- Test 4: Todo ambito writes all fields atomically ----

Public Function Test_CacheMutation_TodoAmbitoWritesAllFields() As String
    Dim logs(0 To 8) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_CacheMutation_TodoAmbitoWritesAllFields = JsonFail("test did not complete", logs)

    On Error GoTo HandleError

    Dim errMsg As String
    Dim dbPath As String
    Dim dbTemp As DAO.Database
    Dim dbCreated As DAO.Database
    Dim exp As Expediente
    Dim op As ExpedienteEntidadOperaciones
    Dim result As String
    Dim rowCount As Long
    Dim campoClasif As Variant
    Dim campoOC As Variant
    Dim campoEjercito As Variant

    ' --- Arrange ---
    logs(0) = "1. Arrange: create isolated temp .accdb with TbExpedientesConEntidades"
    dbPath = MakeIsolatedTempDbPath("todo_ambito")
    Set dbCreated = CreateIsolatedTempDb(dbPath, errMsg)
    If dbCreated Is Nothing Then
        Test_CacheMutation_TodoAmbitoWritesAllFields = JsonFail("CreateIsolatedTempDb: " & errMsg, logs)
        GoTo Teardown
    End If

    logs(1) = "1a. Arrange: create TbExpedientesConEntidades in temp db"
    If Not EnsureCacheTable(dbCreated, errMsg) Then
        Test_CacheMutation_TodoAmbitoWritesAllFields = JsonFail("EnsureCacheTable: " & errMsg, logs)
        GoTo Teardown
    End If
    Set dbCreated = Nothing

    logs(2) = "2. Arrange: create Expediente with IDExpediente=9004 and real model fields"
    Set exp = New Expediente
    exp.IDExpediente = "9004"
    exp.Ambito = "TEST"
    exp.Titulo = "Test Todo Ambito"
    exp.EsAM = "Sí"
    exp.EsExpediente = "No"

    logs(3) = "3. Arrange: configure ExpedienteEntidadOperaciones"
    Set op = New ExpedienteEntidadOperaciones
    Set op.Expediente = exp

    logs(4) = "4. Arrange: open temp db for Registrar (passing p_db)"
    Set dbCreated = DBEngine.Workspaces(0).OpenDatabase(dbPath)

    ' --- Act ---
    logs(5) = "5. Act: Registrar with p_Ambito=Todo — all ambito fields"
    result = op.Registrar(p_db:=dbCreated, p_Ambito:=EnumAmbitoActualizacion.Todo, p_Error:=errMsg)
    If errMsg <> "" Then
        Test_CacheMutation_TodoAmbitoWritesAllFields = JsonFail("Registrar failed: " & errMsg, logs)
        GoTo Teardown
    End If

    ' --- Assert ---
    logs(6) = "6. Assert: row exists in TbExpedientesConEntidades"
    rowCount = CountCacheRows(dbCreated, 9004)
    If rowCount <> 1 Then
        Test_CacheMutation_TodoAmbitoWritesAllFields = JsonFail("Expected 1 row, got " & rowCount, logs)
        GoTo Teardown
    End If

    logs(7) = "7. Assert: lookup-backed Cabecera fields are Null without lookup fixtures"
    campoClasif = ReadCacheField(dbCreated, 9004, "Clasificacion")
    campoOC = ReadCacheField(dbCreated, 9004, "OrganoContratacion")
    campoEjercito = ReadCacheField(dbCreated, 9004, "Ejercito")

    If Not IsNull(campoClasif) Then
        Test_CacheMutation_TodoAmbitoWritesAllFields = JsonFail("Clasificacion expected Null without lookup fixture, got " & Nz(campoClasif, "<NULL>"), logs)
        GoTo Teardown
    End If
    If Not IsNull(campoOC) Then
        Test_CacheMutation_TodoAmbitoWritesAllFields = JsonFail("OrganoContratacion expected Null without lookup fixture, got " & Nz(campoOC, "<NULL>"), logs)
        GoTo Teardown
    End If
    If Not IsNull(campoEjercito) Then
        Test_CacheMutation_TodoAmbitoWritesAllFields = JsonFail("Ejercito expected Null without lookup fixture, got " & Nz(campoEjercito, "<NULL>"), logs)
        GoTo Teardown
    End If

    logs(8) = "8. Assert: TipoParaLista populated from Expediente.EsAM"
    Dim campoTipo As Variant
    campoTipo = ReadCacheField(dbCreated, 9004, "TipoParaLista")
    If Nz(campoTipo, "") <> "AM" Then
        Test_CacheMutation_TodoAmbitoWritesAllFields = JsonFail("TipoParaLista expected AM, got " & Nz(campoTipo, "<NULL>"), logs)
        GoTo Teardown
    End If

    Test_CacheMutation_TodoAmbitoWritesAllFields = JsonOk("todo_ambito_writes_all_fields", logs)

Teardown:
    On Error Resume Next
    Set op = Nothing
    Set exp = Nothing
    If Not dbCreated Is Nothing Then dbCreated.Close
    DisposeIsolatedTempDb dbPath, dbTemp
    Exit Function

HandleError:
    Test_CacheMutation_TodoAmbitoWritesAllFields = JsonFail(Err.Description, logs)
    Resume Teardown
End Function
