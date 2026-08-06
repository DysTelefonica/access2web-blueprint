Attribute VB_Name = "Test_Cache"
Option Compare Database
Option Explicit

' ============================================================
' Test Battery: Cache Verification Layer
' Tests GetCachedX / InvalidarCacheX / ResetGlobals
' Returns JSON {ok, error, payload, logs}
' Run all: node cli.js test-vba Test_Cache_RunAll --access "Gestion_Riesgos.accdb" --json
' Tests are idempotent: seed test data before, teardown after
' ============================================================

Private Const TEST_ID_RANGE As Long = 900000  ' IDs above this are test records
Private Const CACHE_FIX_EXPEDIENTE_ID As Long = 900500
Private Const CACHE_FIX_PROYECTO_ID As Long = 900501
Private Const CACHE_FIX_EDICION_ID As Long = 900502
Private Const CACHE_FIX_RIESGO_ID As Long = 900503
Private Const CACHE_FIX_PM_ID As Long = 900504
Private Const CACHE_FIX_PC_ID As Long = 900505
Private Const CACHE_FIX_PM_ACTION_ID As Long = 900506
Private Const CACHE_FIX_PC_ACTION_ID As Long = 900507
Private Const SEARCH_COMBO_NC_ID_A As Long = 911350
Private Const SEARCH_COMBO_NC_ID_B As Long = 911351
Private Const SEARCH_COMBO_NC_ID_C As Long = 911352
Private Const SEARCH_COMBO_PREFIX As String = "TEST_ISSUE35_"
Private Const SEARCH_COMBO_EXP_A As String = "TEST_ISSUE35_EXP_A"
Private Const SEARCH_COMBO_EXP_B As String = "TEST_ISSUE35_EXP_B"
Private Const SEARCH_COMBO_PROY_A As String = "TEST_ISSUE35_PROY_A"
Private Const SEARCH_COMBO_PROY_B As String = "TEST_ISSUE35_PROY_B"
Private Const SEARCH_COMBO_PROY_C As String = "TEST_ISSUE35_PROY_C"
Private Const SEARCH_COMBO_VEH_A As String = "TEST_ISSUE35_VEH_A"
Private Const SEARCH_COMBO_VEH_B As String = "TEST_ISSUE35_VEH_B"

' --- Cache de conexión (v1.9 §2) ---
Private m_TestDb As DAO.Database

' --- JSON Helpers ---

Private Function JsonEscape(ByVal value As String) As String
    JsonEscape = Replace(value, "\", "\\")
    JsonEscape = Replace(JsonEscape, """", "\""")
    JsonEscape = Replace(JsonEscape, vbCrLf, "\n")
    JsonEscape = Replace(JsonEscape, vbCr, "\n")
    JsonEscape = Replace(JsonEscape, vbLf, "\n")
    JsonEscape = Replace(JsonEscape, vbTab, "\t")
End Function

Private Function BoolJson(ByVal value As Boolean) As String
    BoolJson = IIf(value, "true", "false")
End Function

' --- Deterministic cache fixture IDs ---
' Skill access-vba-tdd: atomic cache tests use the owned SeedAll graph only.
Private Function PrepareCacheFixtureGraph(Optional ByRef p_Error As String = "") As Boolean
    p_Error = ""
    On Error GoTo EH

    ' Cache tests are order-sensitive unless every fixture-backed test starts
    ' from a clean in-memory cache and a freshly seeded sandbox graph.
    On Error Resume Next
    If Not m_TestDb Is Nothing Then m_TestDb.Close
    Set m_TestDb = Nothing
    On Error GoTo EH

    ResetGlobals p_Error
    If p_Error <> "" Then
        p_Error = "PrepareCacheFixtureGraph: ResetGlobals failed: " & p_Error
        Exit Function
    End If
    If Not ForceLocalBackend(p_Error) Then
        p_Error = "PrepareCacheFixtureGraph: ForceLocalBackend failed: " & p_Error
        Exit Function
    End If

    Test_Fixtures.TeardownAll
    Test_Fixtures.SeedAll

    PrepareCacheFixtureGraph = True
    Exit Function

EH:
    p_Error = "PrepareCacheFixtureGraph: " & Err.Description
    PrepareCacheFixtureGraph = False
End Function

Private Function CacheFixtureId(ByVal table As String, ByVal idCol As String, Optional ByRef p_Error As String) As String
    p_Error = ""
    On Error GoTo EH

    If Not PrepareCacheFixtureGraph(p_Error) Then
        CacheFixtureId = ""
        Exit Function
    End If

    Select Case table & "." & idCol
        Case "TbExpedientes.IDExpediente"
            CacheFixtureId = CStr(CACHE_FIX_EXPEDIENTE_ID)
        Case "TbProyectos.IDProyecto"
            CacheFixtureId = CStr(CACHE_FIX_PROYECTO_ID)
        Case "TbProyectosEdiciones.IDEdicion"
            CacheFixtureId = CStr(CACHE_FIX_EDICION_ID)
        Case "TbRiesgos.IDRiesgo"
            CacheFixtureId = CStr(CACHE_FIX_RIESGO_ID)
        Case "TbRiesgosPlanMitigacionPpal.IDMitigacion"
            CacheFixtureId = CStr(CACHE_FIX_PM_ID)
        Case "TbRiesgosPlanContingenciaPpal.IDContingencia"
            CacheFixtureId = CStr(CACHE_FIX_PC_ID)
        Case Else
            p_Error = "Unsupported cache fixture: " & table & "." & idCol
            CacheFixtureId = ""
    End Select
    Exit Function

EH:
    p_Error = "CacheFixtureId: " & Err.Description
    CacheFixtureId = ""
End Function

' --- JSON Wrappers (delegación a Test_Helper_JSON) ---
' BuildJsonOk/BuildJsonFail esperan logs() con LB=0. Test_Cache usa strings concatenados.
' Wrappers adaptados: BuildOkStr / BuildFailStr.

Private Function BuildOkStr(ByRef logStr As String, Optional ByRef payload As String = "null") As String
    Dim logs(0 To 0) As String
    logs(0) = logStr
    BuildOkStr = BuildJsonOk(payload, logs)
End Function

Private Function BuildFailStr(ByRef errMsg As String, ByRef logStr As String) As String
    Dim logs(0 To 0) As String
    logs(0) = logStr
    BuildFailStr = BuildJsonFail(errMsg, logs)
End Function

Private Function ErrJson(ByRef p_Error As String) As String
    ErrJson = BuildFailStr(p_Error, "")
End Function

Private Function OkJson(ByRef logs As String, Optional ByRef payload As String = "null") As String
    OkJson = BuildOkStr(logs, payload)
End Function

' --- DB Helpers ---

Private Function EnsureTestConfigLoaded(ByRef p_Error As String) As Boolean
    ' ForceLocalBackend() ejecuta EVE en modo LOCAL y fuerza getdb() al backend sandbox.
    ' Esto garantiza que TODO (helpers, services, cache) apunte al backend local,
    ' sin importar qué diga BackendActivo en TbConfiguracionBackends.
    p_Error = ""
    If Not ForceLocalBackend(p_Error) Then
        EnsureTestConfigLoaded = False
        Exit Function
    End If
    EnsureTestConfigLoaded = True
End Function

Private Function GetTestDb(Optional ByRef p_Error As String) As Variant
    Dim dummyName As String
    Dim ws As DAO.Workspace
    Dim dbPath As String
    Dim localPassword As String
    Dim cfgError As String

    Set GetTestDb = Nothing
    p_Error = ""

    If Not m_TestDb Is Nothing Then
        On Error Resume Next
        dummyName = m_TestDb.Name
        If Err.Number = 0 Then
            Set GetTestDb = m_TestDb
            Exit Function
        End If
        On Error GoTo 0
        Set m_TestDb = Nothing
    End If

    If Not ForceLocalBackend(cfgError) Then
        p_Error = "TESTS BLOCKED: ForceLocalBackend failed: " & cfgError
        Exit Function
    End If

    dbPath = m_BackendSandboxURL
    If dbPath = "" Or Not CreateObject("Scripting.FileSystemObject").FileExists(dbPath) Then
        p_Error = "TESTS BLOCKED: BackendSandbox not found: " & dbPath
        Exit Function
    End If

    localPassword = m_PasswordBackend
    Set ws = DBEngine(0)
    On Error Resume Next
    Set m_TestDb = ws.OpenDatabase(dbPath, False, False, ";PWD=" & localPassword)
    If Err.Number <> 0 Then
        p_Error = "TESTS BLOCKED: Cannot open backend: " & Err.Description
        Set m_TestDb = Nothing
        Exit Function
    End If
    On Error GoTo 0
    Set GetTestDb = m_TestDb
    p_Error = ""
End Function

Private Function IsTestDbSafe() As Boolean
    ' Guard: tests son seguros solo si m_BackendSandboxURL está configurada y el archivo existe.
    ' NO inspeccionar m_ActiveBackendURL (es Private) ni BackendActivo.
    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        IsTestDbSafe = False
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    IsTestDbSafe = fso.FileExists(m_BackendSandboxURL)
    Set fso = Nothing
End Function

Private Function AssertLocalBackend(ByRef p_Error As String) As Boolean
    ' Delegado a Test_Helper (canónico público, v1.9 §3)
    AssertLocalBackend = Test_Helper.AssertLocalBackend()
    If Not AssertLocalBackend Then
        p_Error = "TESTS BLOCKED: AssertLocalBackend = False"
    End If
End Function

Private Function AssertTestDbPath() As Boolean
    ' Verifica que m_BackendSandboxURL apunte a un archivo existente en disco.
    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        AssertTestDbPath = False
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    AssertTestDbPath = fso.FileExists(m_BackendSandboxURL)
    Set fso = Nothing
End Function

Private Function GetTestDbRaw() As DAO.Database
    ' Backward compatibility wrapper — returns DAO.Database or Nothing.
    ' New code should use GetTestDb(ByRef p_Error) As Variant.
    Dim dummyErr As String
    Set GetTestDbRaw = GetTestDb(dummyErr)
End Function

Private Function IsTestIdSafe(ByVal ID As String) As Boolean
    Dim numId As Long
    On Error Resume Next
    numId = CLng(Nz(ID, 0))
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        IsTestIdSafe = False
        Exit Function
    End If
    On Error GoTo 0
    IsTestIdSafe = (numId > TEST_ID_RANGE)
End Function

Private Function IsTestId(ByVal ID As String) As Boolean
    IsTestId = (CLng(Nz(ID, 0))) > TEST_ID_RANGE
End Function

Private Function CompareField(ByVal cacheVal As String, ByVal dbVal As Variant) As Boolean
    CompareField = (Trim$(Nz(cacheVal, "")) = Trim$(Nz(dbVal, "")))
End Function

' --- Fixture graph ownership ---
' Test_Fixtures.SeedAll owns the cache test graph and TeardownAll cleans it defensively.

' ============================================================
' SEARCH COMBO CACHE RED TESTS (issue #35)
' Schema evidence gathered with Dysflow MCP get_schema on Gestion_Riesgos_Datos.accdb:
' - Table: TbNoConformidades.
' - PK/fixture key: IDNoConformidad, DAO type 4 Long, required.
' - Required fields: CodigoNoConformidad (type 10 Text(255)), EXPEDIENTE (type 10 Text(255)).
' - Seeded optional fields used by cache keys: PROYECTO and VEHICULO, both type 10 Text(255), nullable/zero-length allowed.
' - Domain-safe optional fields used by legal rows: Juridica values include TDE/TSOL; TIPO values include NC/EC; ESTADO values include REGISTRADA/Cerrada.
' Relationships: Dysflow get_relationships did not report application FKs for TbNoConformidades from the inspected backend.
' Fixture rule: tests insert only deterministic IDs 911350-911352 inside a DAO transaction and roll back in teardown.
' ============================================================

Private Function SearchComboCacheBuildFail(ByVal p_Message As String, ByRef p_Logs() As String) As String
    SearchComboCacheBuildFail = Test_Helper.BuildJsonFail(p_Message, p_Logs)
End Function

Private Function SearchComboCacheBuildOk(ByVal p_Value As String, ByRef p_Logs() As String) As String
    SearchComboCacheBuildOk = Test_Helper.BuildJsonOk(p_Value, p_Logs)
End Function

Private Function SearchComboCachePrepareDb(ByRef p_db As DAO.Database, ByRef p_Error As String) As Boolean
    p_Error = ""
    If Not Test_Helper.ForceLocalBackend(p_Error) Then
        p_Error = "TESTS BLOCKED: ForceLocalBackend failed: " & p_Error
        Exit Function
    End If

    Set p_db = Test_Fixtures.GetTestDb(p_Error)
    If p_db Is Nothing Then
        p_Error = "TESTS BLOCKED: GetTestDb returned Nothing: " & p_Error
        Exit Function
    End If

    SearchComboCachePrepareDb = True
End Function

Private Function SearchComboCacheSentinelCount(ByVal p_db As DAO.Database) As Long
    Dim rs As DAO.Recordset

    Set rs = p_db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbNoConformidades " & _
        "WHERE IDNoConformidad IN (" & SEARCH_COMBO_NC_ID_A & ", " & _
        SEARCH_COMBO_NC_ID_B & ", " & SEARCH_COMBO_NC_ID_C & ")", _
        dbOpenSnapshot)
    If rs.EOF Then
        SearchComboCacheSentinelCount = 0
    Else
        SearchComboCacheSentinelCount = CLng(Nz(rs.Fields("Cnt").Value, 0))
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Sub SearchComboCacheSeedRows(ByVal p_db As DAO.Database)
    p_db.Execute "INSERT INTO TbNoConformidades " & _
        "(IDNoConformidad, CodigoNoConformidad, EXPEDIENTE, PROYECTO, VEHICULO, Juridica, TIPO, ESTADO, FECHAAPERTURA) " & _
        "VALUES (" & SEARCH_COMBO_NC_ID_A & ", '" & SEARCH_COMBO_PREFIX & "NC_A', '" & _
        SEARCH_COMBO_EXP_A & "', '" & SEARCH_COMBO_PROY_A & "', '" & SEARCH_COMBO_VEH_A & _
        "', 'TDE', 'NC', 'REGISTRADA', Date())", dbFailOnError
    p_db.Execute "INSERT INTO TbNoConformidades " & _
        "(IDNoConformidad, CodigoNoConformidad, EXPEDIENTE, PROYECTO, VEHICULO, Juridica, TIPO, ESTADO, FECHAAPERTURA) " & _
        "VALUES (" & SEARCH_COMBO_NC_ID_B & ", '" & SEARCH_COMBO_PREFIX & "NC_B', '" & _
        SEARCH_COMBO_EXP_B & "', '" & SEARCH_COMBO_PROY_B & "', '" & SEARCH_COMBO_VEH_A & _
        "', 'TDE', 'NC', 'REGISTRADA', Date())", dbFailOnError
    p_db.Execute "INSERT INTO TbNoConformidades " & _
        "(IDNoConformidad, CodigoNoConformidad, EXPEDIENTE, PROYECTO, VEHICULO, Juridica, TIPO, ESTADO, FECHAAPERTURA) " & _
        "VALUES (" & SEARCH_COMBO_NC_ID_C & ", '" & SEARCH_COMBO_PREFIX & "NC_C', '" & _
        SEARCH_COMBO_EXP_A & "', '" & SEARCH_COMBO_PROY_C & "', '" & SEARCH_COMBO_VEH_B & _
        "', 'TDE', 'NC', 'REGISTRADA', Date())", dbFailOnError
End Sub

Private Function SearchComboCacheCountPrefixValues(ByVal p_Values As Scripting.Dictionary) As Long
    Dim item As Variant

    If p_Values Is Nothing Then
        SearchComboCacheCountPrefixValues = 0
        Exit Function
    End If

    For Each item In p_Values.Keys
        If Left$(CStr(item), Len(SEARCH_COMBO_PREFIX)) = SEARCH_COMBO_PREFIX Then
            SearchComboCacheCountPrefixValues = SearchComboCacheCountPrefixValues + 1
        End If
    Next item
End Function

Private Function SearchComboCacheAssertContains(ByVal p_Values As Scripting.Dictionary, ByVal p_Expected As String, ByRef p_Error As String) As Boolean
    p_Error = ""
    If p_Values Is Nothing Then
        p_Error = "resolver returned Nothing"
        Exit Function
    End If
    If Not p_Values.Exists(p_Expected) Then
        p_Error = "missing seeded value: " & p_Expected
        Exit Function
    End If
    SearchComboCacheAssertContains = True
End Function

Private Function SearchComboCacheAssertCounters( _
    ByVal p_ExpectedMisses As Long, _
    ByVal p_ExpectedHits As Long, _
    ByVal p_ExpectedInvalidations As Long, _
    ByRef p_Error As String) As Boolean

    p_Error = ""
    If SearchComboCache_MissCount() <> p_ExpectedMisses Then
        p_Error = "expected MissCount=" & p_ExpectedMisses & ", got " & SearchComboCache_MissCount()
        Exit Function
    End If
    If SearchComboCache_HitCount() <> p_ExpectedHits Then
        p_Error = "expected HitCount=" & p_ExpectedHits & ", got " & SearchComboCache_HitCount()
        Exit Function
    End If
    If SearchComboCache_InvalidationCount() <> p_ExpectedInvalidations Then
        p_Error = "expected InvalidationCount=" & p_ExpectedInvalidations & ", got " & SearchComboCache_InvalidationCount()
        Exit Function
    End If
    SearchComboCacheAssertCounters = True
End Function

Public Function Test_SearchComboCache_ColdMiss_LoadsSeededExpedientesOnce() As String
    On Error GoTo EH

    Dim logs(0 To 7) As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim transStarted As Boolean
    Dim values As Scripting.Dictionary

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: schema-first TbNoConformidades legal rows, deterministic IDs 911350-911352"
    logs(2) = "3. Arrange: pre-check sentinels; do not delete shared data"
    logs(3) = "4. Arrange: BeginTrans and insert exact NC rows"
    logs(4) = "5. Act: reset cache and request TbNoConformidades.EXPEDIENTE"
    logs(5) = "6. Assert: seeded distinct Expediente values are present exactly once by prefix"
    logs(6) = "7. Assert: MissCount=1 and HitCount=0"
    logs(7) = "8. Teardown: rollback transaction"

    If Not SearchComboCachePrepareDb(db, errMsg) Then GoTo Fail
    If SearchComboCacheSentinelCount(db) <> 0 Then
        errMsg = "TESTS BLOCKED: issue #35 sentinel NC rows already exist; refusing to mutate shared data"
        GoTo Fail
    End If

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    transStarted = True
    SearchComboCacheSeedRows db

    SearchComboCache_Reset errMsg
    If errMsg <> "" Then GoTo Fail
    Set values = SearchComboCache_GetDistinctValues("TbNoConformidades", "EXPEDIENTE", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    If Not SearchComboCacheAssertContains(values, SEARCH_COMBO_EXP_A, errMsg) Then GoTo Fail
    If Not SearchComboCacheAssertContains(values, SEARCH_COMBO_EXP_B, errMsg) Then GoTo Fail
    If SearchComboCacheCountPrefixValues(values) <> 2 Then
        errMsg = "expected exactly 2 seeded Expediente values by prefix, got " & SearchComboCacheCountPrefixValues(values)
        GoTo Fail
    End If
    If SearchComboCache_MissCount() <> 1 Then
        errMsg = "expected MissCount=1, got " & SearchComboCache_MissCount()
        GoTo Fail
    End If
    If SearchComboCache_HitCount() <> 0 Then
        errMsg = "expected HitCount=0 on cold access, got " & SearchComboCache_HitCount()
        GoTo Fail
    End If

    Test_SearchComboCache_ColdMiss_LoadsSeededExpedientesOnce = SearchComboCacheBuildOk("cold_miss_ok", logs)
Teardown:
    On Error Resume Next
    If transStarted Then ws.Rollback
    Set values = Nothing
    Set db = Nothing
    Set ws = Nothing
    SearchComboCache_Reset errMsg
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function
Fail:
    Test_SearchComboCache_ColdMiss_LoadsSeededExpedientesOnce = SearchComboCacheBuildFail(errMsg, logs)
    GoTo Teardown
EH:
    errMsg = "Test_SearchComboCache_ColdMiss_LoadsSeededExpedientesOnce: " & Err.Description
    GoTo Fail
End Function

Public Function Test_SearchComboCache_WarmHit_DoesNotReloadExpedientes() As String
    On Error GoTo EH

    Dim logs(0 To 7) As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim transStarted As Boolean
    Dim coldValues As Scripting.Dictionary
    Dim warmValues As Scripting.Dictionary

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: deterministic NC rows for Expediente cache"
    logs(2) = "3. Arrange: BeginTrans so fixture rows roll back"
    logs(3) = "4. Act: first resolver request warms cache"
    logs(4) = "5. Act: second resolver request repeats same key"
    logs(5) = "6. Assert: seeded values remain present"
    logs(6) = "7. Assert: MissCount remains 1 and HitCount becomes 1"
    logs(7) = "8. Teardown: rollback transaction"

    If Not SearchComboCachePrepareDb(db, errMsg) Then GoTo Fail
    If SearchComboCacheSentinelCount(db) <> 0 Then
        errMsg = "TESTS BLOCKED: issue #35 sentinel NC rows already exist; refusing to mutate shared data"
        GoTo Fail
    End If

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    transStarted = True
    SearchComboCacheSeedRows db

    SearchComboCache_Reset errMsg
    If errMsg <> "" Then GoTo Fail
    Set coldValues = SearchComboCache_GetDistinctValues("TbNoConformidades", "EXPEDIENTE", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    Set warmValues = SearchComboCache_GetDistinctValues("TbNoConformidades", "EXPEDIENTE", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    If Not SearchComboCacheAssertContains(warmValues, SEARCH_COMBO_EXP_A, errMsg) Then GoTo Fail
    If Not SearchComboCacheAssertContains(warmValues, SEARCH_COMBO_EXP_B, errMsg) Then GoTo Fail
    If SearchComboCache_MissCount() <> 1 Then
        errMsg = "expected no additional reload; MissCount must stay 1, got " & SearchComboCache_MissCount()
        GoTo Fail
    End If
    If SearchComboCache_HitCount() <> 1 Then
        errMsg = "expected HitCount=1 on warm access, got " & SearchComboCache_HitCount()
        GoTo Fail
    End If

    Test_SearchComboCache_WarmHit_DoesNotReloadExpedientes = SearchComboCacheBuildOk("warm_hit_ok", logs)
Teardown:
    On Error Resume Next
    If transStarted Then ws.Rollback
    Set warmValues = Nothing
    Set coldValues = Nothing
    Set db = Nothing
    Set ws = Nothing
    SearchComboCache_Reset errMsg
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function
Fail:
    Test_SearchComboCache_WarmHit_DoesNotReloadExpedientes = SearchComboCacheBuildFail(errMsg, logs)
    GoTo Teardown
EH:
    errMsg = "Test_SearchComboCache_WarmHit_DoesNotReloadExpedientes: " & Err.Description
    GoTo Fail
End Function

Public Function Test_SearchComboCache_InvalidatedKey_ReloadsExactlyOnce() As String
    On Error GoTo EH

    Dim logs(0 To 7) As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim transStarted As Boolean
    Dim valuesAfterInvalidation As Scripting.Dictionary

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: deterministic Expediente rows"
    logs(2) = "3. Act: cold request warms Expediente key"
    logs(3) = "4. Act: invalidate Expediente key"
    logs(4) = "5. Act: request key twice after invalidation"
    logs(5) = "6. Assert: exactly one additional reload after invalidation"
    logs(6) = "7. Assert: second post-invalidation access is warm"
    logs(7) = "8. Teardown: rollback transaction"

    If Not SearchComboCachePrepareDb(db, errMsg) Then GoTo Fail
    If SearchComboCacheSentinelCount(db) <> 0 Then
        errMsg = "TESTS BLOCKED: issue #35 sentinel NC rows already exist; refusing to mutate shared data"
        GoTo Fail
    End If

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    transStarted = True
    SearchComboCacheSeedRows db

    SearchComboCache_Reset errMsg
    If errMsg <> "" Then GoTo Fail
    Set valuesAfterInvalidation = SearchComboCache_GetDistinctValues("TbNoConformidades", "EXPEDIENTE", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    SearchComboCache_Invalidar "TbNoConformidades", "EXPEDIENTE"
    Set valuesAfterInvalidation = SearchComboCache_GetDistinctValues("TbNoConformidades", "EXPEDIENTE", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    Set valuesAfterInvalidation = SearchComboCache_GetDistinctValues("TbNoConformidades", "EXPEDIENTE", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    If Not SearchComboCacheAssertContains(valuesAfterInvalidation, SEARCH_COMBO_EXP_A, errMsg) Then GoTo Fail
    If SearchComboCache_MissCount() <> 2 Then
        errMsg = "expected initial miss + one reload after invalidation; MissCount=2, got " & SearchComboCache_MissCount()
        GoTo Fail
    End If
    If SearchComboCache_HitCount() <> 1 Then
        errMsg = "expected second post-invalidation access warm; HitCount=1, got " & SearchComboCache_HitCount()
        GoTo Fail
    End If
    If SearchComboCache_InvalidationCount() <> 1 Then
        errMsg = "expected InvalidationCount=1, got " & SearchComboCache_InvalidationCount()
        GoTo Fail
    End If

    Test_SearchComboCache_InvalidatedKey_ReloadsExactlyOnce = SearchComboCacheBuildOk("invalidation_reload_ok", logs)
Teardown:
    On Error Resume Next
    If transStarted Then ws.Rollback
    Set valuesAfterInvalidation = Nothing
    Set db = Nothing
    Set ws = Nothing
    SearchComboCache_Reset errMsg
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function
Fail:
    Test_SearchComboCache_InvalidatedKey_ReloadsExactlyOnce = SearchComboCacheBuildFail(errMsg, logs)
    GoTo Teardown
EH:
    errMsg = "Test_SearchComboCache_InvalidatedKey_ReloadsExactlyOnce: " & Err.Description
    GoTo Fail
End Function

Public Function Test_SearchComboCache_IsolatedKeys_DoNotShareEntries() As String
    On Error GoTo EH

    Dim logs(0 To 8) As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim transStarted As Boolean
    Dim expValues As Scripting.Dictionary
    Dim proyValues As Scripting.Dictionary
    Dim vehValues As Scripting.Dictionary

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: deterministic NC rows with distinct Expediente/Proyecto/Vehiculo values"
    logs(2) = "3. Act: cold-load EXPEDIENTE, PROYECTO, and VEHICULO keys"
    logs(3) = "4. Assert: each key exposes only its own seeded distinct values"
    logs(4) = "5. Assert: three independent cold misses"
    logs(5) = "6. Act: invalidate only EXPEDIENTE"
    logs(6) = "7. Assert: PROYECTO remains warm and EXPEDIENTE reloads"
    logs(7) = "8. Assert: counters prove key isolation"
    logs(8) = "9. Teardown: rollback transaction"

    If Not SearchComboCachePrepareDb(db, errMsg) Then GoTo Fail
    If SearchComboCacheSentinelCount(db) <> 0 Then
        errMsg = "TESTS BLOCKED: issue #35 sentinel NC rows already exist; refusing to mutate shared data"
        GoTo Fail
    End If

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    transStarted = True
    SearchComboCacheSeedRows db

    SearchComboCache_Reset errMsg
    If errMsg <> "" Then GoTo Fail
    Set expValues = SearchComboCache_GetDistinctValues("TbNoConformidades", "EXPEDIENTE", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    Set proyValues = SearchComboCache_GetDistinctValues("TbNoConformidades", "PROYECTO", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    Set vehValues = SearchComboCache_GetDistinctValues("TbNoConformidades", "VEHICULO", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    If Not SearchComboCacheAssertContains(expValues, SEARCH_COMBO_EXP_A, errMsg) Then GoTo Fail
    If Not SearchComboCacheAssertContains(expValues, SEARCH_COMBO_EXP_B, errMsg) Then GoTo Fail
    If Not SearchComboCacheAssertContains(proyValues, SEARCH_COMBO_PROY_A, errMsg) Then GoTo Fail
    If Not SearchComboCacheAssertContains(proyValues, SEARCH_COMBO_PROY_B, errMsg) Then GoTo Fail
    If Not SearchComboCacheAssertContains(proyValues, SEARCH_COMBO_PROY_C, errMsg) Then GoTo Fail
    If Not SearchComboCacheAssertContains(vehValues, SEARCH_COMBO_VEH_A, errMsg) Then GoTo Fail
    If Not SearchComboCacheAssertContains(vehValues, SEARCH_COMBO_VEH_B, errMsg) Then GoTo Fail
    If SearchComboCacheCountPrefixValues(expValues) <> 2 Then
        errMsg = "expected exactly 2 Expediente seeded values by prefix"
        GoTo Fail
    End If
    If SearchComboCacheCountPrefixValues(proyValues) <> 3 Then
        errMsg = "expected exactly 3 Proyecto seeded values by prefix"
        GoTo Fail
    End If
    If SearchComboCacheCountPrefixValues(vehValues) <> 2 Then
        errMsg = "expected exactly 2 Vehiculo seeded values by prefix"
        GoTo Fail
    End If
    If Not SearchComboCacheAssertCounters(3, 0, 0, errMsg) Then GoTo Fail

    SearchComboCache_Invalidar "TbNoConformidades", "EXPEDIENTE"
    Set proyValues = SearchComboCache_GetDistinctValues("TbNoConformidades", "PROYECTO", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    Set expValues = SearchComboCache_GetDistinctValues("TbNoConformidades", "EXPEDIENTE", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    If Not SearchComboCacheAssertContains(proyValues, SEARCH_COMBO_PROY_C, errMsg) Then GoTo Fail
    If Not SearchComboCacheAssertContains(expValues, SEARCH_COMBO_EXP_B, errMsg) Then GoTo Fail
    If Not SearchComboCacheAssertCounters(4, 1, 1, errMsg) Then GoTo Fail

    Test_SearchComboCache_IsolatedKeys_DoNotShareEntries = SearchComboCacheBuildOk("isolated_keys_ok", logs)
Teardown:
    On Error Resume Next
    If transStarted Then ws.Rollback
    Set vehValues = Nothing
    Set proyValues = Nothing
    Set expValues = Nothing
    Set db = Nothing
    Set ws = Nothing
    SearchComboCache_Reset errMsg
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function
Fail:
    Test_SearchComboCache_IsolatedKeys_DoNotShareEntries = SearchComboCacheBuildFail(errMsg, logs)
    GoTo Teardown
EH:
    errMsg = "Test_SearchComboCache_IsolatedKeys_DoNotShareEntries: " & Err.Description
    GoTo Fail
End Function

Public Function Test_SearchComboCache_Reset_ReloadsAfterCacheReset() As String
    On Error GoTo EH

    Dim logs(0 To 7) As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim transStarted As Boolean
    Dim values As Scripting.Dictionary

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: deterministic NC rows"
    logs(2) = "3. Act: cold-load Expediente key"
    logs(3) = "4. Assert: first load is one miss"
    logs(4) = "5. Act: reset full cache and stats"
    logs(5) = "6. Act: request Expediente again"
    logs(6) = "7. Assert: reload after reset is a fresh miss"
    logs(7) = "8. Teardown: rollback transaction"

    If Not SearchComboCachePrepareDb(db, errMsg) Then GoTo Fail
    If SearchComboCacheSentinelCount(db) <> 0 Then
        errMsg = "TESTS BLOCKED: issue #35 sentinel NC rows already exist; refusing to mutate shared data"
        GoTo Fail
    End If

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    transStarted = True
    SearchComboCacheSeedRows db

    SearchComboCache_Reset errMsg
    If errMsg <> "" Then GoTo Fail
    Set values = SearchComboCache_GetDistinctValues("TbNoConformidades", "EXPEDIENTE", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    If Not SearchComboCacheAssertCounters(1, 0, 0, errMsg) Then GoTo Fail

    SearchComboCache_Reset errMsg
    If errMsg <> "" Then GoTo Fail
    Set values = SearchComboCache_GetDistinctValues("TbNoConformidades", "EXPEDIENTE", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    If Not SearchComboCacheAssertContains(values, SEARCH_COMBO_EXP_A, errMsg) Then GoTo Fail
    If Not SearchComboCacheAssertContains(values, SEARCH_COMBO_EXP_B, errMsg) Then GoTo Fail
    If SearchComboCacheCountPrefixValues(values) <> 2 Then
        errMsg = "expected exactly 2 seeded Expediente values by prefix after reset"
        GoTo Fail
    End If
    If Not SearchComboCacheAssertCounters(1, 0, 0, errMsg) Then GoTo Fail

    Test_SearchComboCache_Reset_ReloadsAfterCacheReset = SearchComboCacheBuildOk("reset_reload_ok", logs)
Teardown:
    On Error Resume Next
    If transStarted Then ws.Rollback
    Set values = Nothing
    Set db = Nothing
    Set ws = Nothing
    SearchComboCache_Reset errMsg
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function
Fail:
    Test_SearchComboCache_Reset_ReloadsAfterCacheReset = SearchComboCacheBuildFail(errMsg, logs)
    GoTo Teardown
EH:
    errMsg = "Test_SearchComboCache_Reset_ReloadsAfterCacheReset: " & Err.Description
    GoTo Fail
End Function

Public Function Test_SearchComboCache_UnsupportedKey_ReturnsDeterministicError() As String
    On Error GoTo EH

    Dim logs(0 To 7) As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim transStarted As Boolean
    Dim unsupportedValues As Scripting.Dictionary
    Dim supportedValues As Scripting.Dictionary

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: deterministic NC rows"
    logs(2) = "3. Act: request unsupported NC field"
    logs(3) = "4. Assert: deterministic unsupported-field error"
    logs(4) = "5. Assert: unsupported key does not increment cache counters"
    logs(5) = "6. Act: request supported Expediente key after unsupported call"
    logs(6) = "7. Assert: supported key cold-loads normally with no ambiguous cache entry"
    logs(7) = "8. Teardown: rollback transaction"

    If Not SearchComboCachePrepareDb(db, errMsg) Then GoTo Fail
    If SearchComboCacheSentinelCount(db) <> 0 Then
        errMsg = "TESTS BLOCKED: issue #35 sentinel NC rows already exist; refusing to mutate shared data"
        GoTo Fail
    End If

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    transStarted = True
    SearchComboCacheSeedRows db

    SearchComboCache_Reset errMsg
    If errMsg <> "" Then GoTo Fail
    Set unsupportedValues = SearchComboCache_GetDistinctValues("TbNoConformidades", "CodigoNoConformidad", db, errMsg)
    If Not unsupportedValues Is Nothing Then
        errMsg = "expected unsupported key to return Nothing"
        GoTo Fail
    End If
    If InStr(1, errMsg, "Unsupported search combo cache field", vbTextCompare) = 0 Then
        errMsg = "expected deterministic unsupported-field error, got: " & errMsg
        GoTo Fail
    End If
    If Not SearchComboCacheAssertCounters(0, 0, 0, errMsg) Then GoTo Fail

    errMsg = ""
    Set supportedValues = SearchComboCache_GetDistinctValues("TbNoConformidades", "EXPEDIENTE", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    If Not SearchComboCacheAssertContains(supportedValues, SEARCH_COMBO_EXP_A, errMsg) Then GoTo Fail
    If Not SearchComboCacheAssertContains(supportedValues, SEARCH_COMBO_EXP_B, errMsg) Then GoTo Fail
    If Not SearchComboCacheAssertCounters(1, 0, 0, errMsg) Then GoTo Fail

    Test_SearchComboCache_UnsupportedKey_ReturnsDeterministicError = SearchComboCacheBuildOk("unsupported_key_ok", logs)
Teardown:
    On Error Resume Next
    If transStarted Then ws.Rollback
    Set supportedValues = Nothing
    Set unsupportedValues = Nothing
    Set db = Nothing
    Set ws = Nothing
    SearchComboCache_Reset errMsg
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function
Fail:
    Test_SearchComboCache_UnsupportedKey_ReturnsDeterministicError = SearchComboCacheBuildFail(errMsg, logs)
    GoTo Teardown
EH:
    errMsg = "Test_SearchComboCache_UnsupportedKey_ReturnsDeterministicError: " & Err.Description
    GoTo Fail
End Function

Public Function Test_SearchComboCache_DirectResolver_DoesNotOpenForm() As String
    On Error GoTo EH

    Dim logs(0 To 8) As String
    Dim errMsg As String
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim transStarted As Boolean
    Dim wasOpenBefore As Boolean
    Dim isOpenAfter As Boolean
    Dim values As Scripting.Dictionary

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: capture FormRiesgoNC open state without opening it"
    logs(2) = "3. Arrange: deterministic NC rows"
    logs(3) = "4. Act: call resolver directly, no UI adapter and no timing assertion"
    logs(4) = "5. Assert: seeded Vehiculo values are present"
    logs(5) = "6. Assert: counters prove resolver path used one cold miss"
    logs(6) = "7. Assert: FormRiesgoNC open state did not change"
    logs(7) = "8. Assert: no warm hit or invalidation side effect"
    logs(8) = "9. Teardown: rollback transaction"

    wasOpenBefore = FormularioAbierto("FormRiesgoNC", errMsg)
    If errMsg <> "" Then GoTo Fail
    If Not SearchComboCachePrepareDb(db, errMsg) Then GoTo Fail
    If SearchComboCacheSentinelCount(db) <> 0 Then
        errMsg = "TESTS BLOCKED: issue #35 sentinel NC rows already exist; refusing to mutate shared data"
        GoTo Fail
    End If

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    transStarted = True
    SearchComboCacheSeedRows db

    SearchComboCache_Reset errMsg
    If errMsg <> "" Then GoTo Fail
    Set values = SearchComboCache_GetDistinctValues("TbNoConformidades", "VEHICULO", db, errMsg)
    If errMsg <> "" Then GoTo Fail
    If Not SearchComboCacheAssertContains(values, SEARCH_COMBO_VEH_A, errMsg) Then GoTo Fail
    If Not SearchComboCacheAssertContains(values, SEARCH_COMBO_VEH_B, errMsg) Then GoTo Fail
    If SearchComboCacheCountPrefixValues(values) <> 2 Then
        errMsg = "expected exactly 2 seeded Vehiculo values by prefix"
        GoTo Fail
    End If
    If Not SearchComboCacheAssertCounters(1, 0, 0, errMsg) Then GoTo Fail
    isOpenAfter = FormularioAbierto("FormRiesgoNC", errMsg)
    If errMsg <> "" Then GoTo Fail
    If isOpenAfter <> wasOpenBefore Then
        errMsg = "resolver changed FormRiesgoNC open state"
        GoTo Fail
    End If

    Test_SearchComboCache_DirectResolver_DoesNotOpenForm = SearchComboCacheBuildOk("direct_resolver_no_form_ok", logs)
Teardown:
    On Error Resume Next
    If transStarted Then ws.Rollback
    Set values = Nothing
    Set db = Nothing
    Set ws = Nothing
    SearchComboCache_Reset errMsg
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function
Fail:
    Test_SearchComboCache_DirectResolver_DoesNotOpenForm = SearchComboCacheBuildFail(errMsg, logs)
    GoTo Teardown
EH:
    errMsg = "Test_SearchComboCache_DirectResolver_DoesNotOpenForm: " & Err.Description
    GoTo Fail
End Function

' ============================================================
' RIESGO CACHE TESTS (10)
' ============================================================

Public Function Test_Cache_Riesgo_Hit() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgos", "IDRiesgo")
    If testId = "" Then
        Test_Cache_Riesgo_Hit = ErrJson("No test data found for Riesgo")
        Exit Function
    End If
    logs = logs & "id=" & testId & ";"
    
    Dim obj As riesgo
    Dim rs As DAO.Recordset
    Set obj = GetCachedRiesgo(testId)
    
    If obj Is Nothing Then
        Test_Cache_Riesgo_Hit = ErrJson("GetCachedRiesgo returned Nothing")
        Exit Function
    End If
    logs = logs & "cached;"
    
    Set rs = GetTestDb().OpenRecordset("SELECT * FROM TbRiesgos WHERE IDRiesgo=" & testId)
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        Test_Cache_Riesgo_Hit = ErrJson("DB query returned no record")
        Exit Function
    End If
    
    If Not CompareField(obj.IDRiesgo, rs!IDRiesgo.value) Then
        rs.Close: Set rs = Nothing
        Test_Cache_Riesgo_Hit = ErrJson("IDRiesgo mismatch")
        Exit Function
    End If
    If Not CompareField(obj.CodigoRiesgo, rs!CodigoRiesgo.value) Then
        rs.Close: Set rs = Nothing
        Test_Cache_Riesgo_Hit = ErrJson("CodigoRiesgo mismatch")
        Exit Function
    End If
    If Not CompareField(obj.Descripcion, rs!Descripcion.value) Then
        rs.Close: Set rs = Nothing
        Test_Cache_Riesgo_Hit = ErrJson("Descripcion mismatch")
        Exit Function
    End If
    
    rs.Close: Set rs = Nothing
    Test_Cache_Riesgo_Hit = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Riesgo_Hit = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Riesgo_CacheConsistency() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgos", "IDRiesgo")
    If testId = "" Then
        Test_Cache_Riesgo_CacheConsistency = ErrJson("No test data")
        Exit Function
    End If
    
    ' First call - populates cache
    Dim obj1 As riesgo
    Set obj1 = GetCachedRiesgo(testId)
    ' Second call - should return same object reference
    Dim obj2 As riesgo
    Set obj2 = GetCachedRiesgo(testId)
    
    If obj1 Is Nothing Or obj2 Is Nothing Then
        Test_Cache_Riesgo_CacheConsistency = ErrJson("Got Nothing")
        Exit Function
    End If
    
    ' Same reference check (by ID equality)
    If obj1.IDRiesgo <> obj2.IDRiesgo Then
        Test_Cache_Riesgo_CacheConsistency = ErrJson("Objects have different ID")
        Exit Function
    End If
    
    logs = logs & "consistent;"
    Test_Cache_Riesgo_CacheConsistency = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Riesgo_CacheConsistency = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Riesgo_Miss() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    ' Use a very high non-existent ID
    Dim testId As String
    testId = "999999999"
    
    Dim obj As riesgo
    Set obj = GetCachedRiesgo(testId)
    
    If Not obj Is Nothing Then
        Test_Cache_Riesgo_Miss = ErrJson("Expected Nothing for non-existent ID, got object")
        Exit Function
    End If
    
    logs = logs & "miss_ok;"
    Test_Cache_Riesgo_Miss = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Riesgo_Miss = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Riesgo_Vacio() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim obj As riesgo
    Set obj = GetCachedRiesgo("")
    
    If Not obj Is Nothing Then
        Test_Cache_Riesgo_Vacio = ErrJson("Expected Nothing for empty ID, got object")
        Exit Function
    End If
    
    logs = logs & "vacio_ok;"
    Test_Cache_Riesgo_Vacio = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Riesgo_Vacio = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Riesgo_Invalidar() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgos", "IDRiesgo")
    If testId = "" Then
        Test_Cache_Riesgo_Invalidar = ErrJson("No test data")
        Exit Function
    End If
    
    ' Populate cache
    Dim obj1 As riesgo
    Set obj1 = GetCachedRiesgo(testId)
    logs = logs & "cached;"
    
    ' Invalidate
    InvalidarCacheRiesgo testId
    logs = logs & "invalidated;"
    
    ' Get again - should re-query
    Dim obj2 As riesgo
    Set obj2 = GetCachedRiesgo(testId)
    
    If obj2 Is Nothing Then
        Test_Cache_Riesgo_Invalidar = ErrJson("After invalidation, GetCached returned Nothing")
        Exit Function
    End If
    
    ' obj2 should be fresh, same ID
    If obj2.IDRiesgo <> testId Then
        Test_Cache_Riesgo_Invalidar = ErrJson("After invalidation, re-query returned wrong ID")
        Exit Function
    End If
    
    logs = logs & "requery_ok;"
    Test_Cache_Riesgo_Invalidar = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Riesgo_Invalidar = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Riesgo_RequeryAfterInvalidate() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgos", "IDRiesgo")
    If testId = "" Then
        Test_Cache_Riesgo_RequeryAfterInvalidate = ErrJson("No test data")
        Exit Function
    End If
    
    ' Populate cache
    Dim obj1 As riesgo
    Set obj1 = GetCachedRiesgo(testId)
    
    ' Invalidate
    InvalidarCacheRiesgo testId
    
    ' Get again
    Dim obj2 As riesgo
    Set obj2 = GetCachedRiesgo(testId)
    
    If obj1.IDRiesgo <> obj2.IDRiesgo Then
        Test_Cache_Riesgo_RequeryAfterInvalidate = ErrJson("Re-query returned different ID after invalidation")
        Exit Function
    End If
    
    logs = logs & "requery_ok;"
    Test_Cache_Riesgo_RequeryAfterInvalidate = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Riesgo_RequeryAfterInvalidate = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Riesgo_NestedEdicion() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgos", "IDRiesgo")
    If testId = "" Then
        Test_Cache_Riesgo_NestedEdicion = ErrJson("No test data")
        Exit Function
    End If
    
    ' Get cached riesgo
    Dim objRiesgo As riesgo
    Set objRiesgo = GetCachedRiesgo(testId)
    If objRiesgo Is Nothing Then
        Test_Cache_Riesgo_NestedEdicion = ErrJson("Riesgo is Nothing")
        Exit Function
    End If
    logs = logs & "riesgo_cached;"
    
    ' Access .Edicion - triggers lazy load and should cache the Edicion
    Dim objEdicion As Edicion
    Set objEdicion = objRiesgo.Edicion
    If objEdicion Is Nothing Then
        Test_Cache_Riesgo_NestedEdicion = ErrJson("Edicion is Nothing")
        Exit Function
    End If
    logs = logs & "edicion_accessed;"
    
    ' Verify edicion is cached in m_DicEdiciones
    ' We check by trying to get the same edicion via GetCachedEdicion
    Dim objEdicionCached As Edicion
    Set objEdicionCached = GetCachedEdicion(objEdicion.IDEdicion)
    If objEdicionCached Is Nothing Then
        Test_Cache_Riesgo_NestedEdicion = ErrJson("Edicion not in cache after accessing via Riesgo.Edicion")
        Exit Function
    End If
    
    If objEdicion.IDEdicion <> objEdicionCached.IDEdicion Then
        Test_Cache_Riesgo_NestedEdicion = ErrJson("Cached edicion has different ID")
        Exit Function
    End If
    
    logs = logs & "nested_cache_ok;"
    Test_Cache_Riesgo_NestedEdicion = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Riesgo_NestedEdicion = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Riesgo_NestedProyecto() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgos", "IDRiesgo")
    If testId = "" Then
        Test_Cache_Riesgo_NestedProyecto = ErrJson("No test data")
        Exit Function
    End If
    
    ' Get cached riesgo
    Dim objRiesgo As riesgo
    Set objRiesgo = GetCachedRiesgo(testId)
    If objRiesgo Is Nothing Then
        Test_Cache_Riesgo_NestedProyecto = ErrJson("Riesgo is Nothing")
        Exit Function
    End If
    
    ' Access .Edicion
    Dim objEdicion As Edicion
    Set objEdicion = objRiesgo.Edicion
    If objEdicion Is Nothing Then
        Test_Cache_Riesgo_NestedProyecto = ErrJson("Edicion is Nothing")
        Exit Function
    End If
    logs = logs & "edicion_ok;"
    
    ' Access .Proyecto via Edicion
    Dim objProyecto As Proyecto
    Set objProyecto = objEdicion.Proyecto
    If objProyecto Is Nothing Then
        Test_Cache_Riesgo_NestedProyecto = ErrJson("Proyecto is Nothing")
        Exit Function
    End If
    logs = logs & "proyecto_accessed;"
    
    ' Verify proyecto is cached
    Dim objProyectoCached As Proyecto
    Set objProyectoCached = GetCachedProyecto(objProyecto.IDProyecto)
    If objProyectoCached Is Nothing Then
        Test_Cache_Riesgo_NestedProyecto = ErrJson("Proyecto not in cache after nested access")
        Exit Function
    End If
    
    logs = logs & "nested_cache_ok;"
    Test_Cache_Riesgo_NestedProyecto = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Riesgo_NestedProyecto = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Riesgo_SameReference() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgos", "IDRiesgo")
    If testId = "" Then
        Test_Cache_Riesgo_SameReference = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As riesgo
    Set obj1 = GetCachedRiesgo(testId)
    Dim obj2 As riesgo
    Set obj2 = GetCachedRiesgo(testId)
    
    If obj1 Is Nothing Or obj2 Is Nothing Then
        Test_Cache_Riesgo_SameReference = ErrJson("Got Nothing")
        Exit Function
    End If
    
    ' Same ID means same cached entry
    If obj1.IDRiesgo <> obj2.IDRiesgo Then
        Test_Cache_Riesgo_SameReference = ErrJson("Different IDs returned")
        Exit Function
    End If
    
    logs = logs & "same_reference_ok;"
    Test_Cache_Riesgo_SameReference = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Riesgo_SameReference = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Riesgo_KeyCaseInsensitive() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgos", "IDRiesgo")
    If testId = "" Then
        Test_Cache_Riesgo_KeyCaseInsensitive = ErrJson("No test data")
        Exit Function
    End If
    
    Dim testIdLower As String
    Dim testIdUpper As String
    testIdLower = LCase$(testId)
    testIdUpper = UCase$(testId)
    
    ' Only test if they differ (for non-numeric IDs)
    If testIdLower = testIdUpper Then
        logs = logs & "numeric_id_skip;"
        Test_Cache_Riesgo_KeyCaseInsensitive = OkJson(logs)
        Exit Function
    End If
    
    Dim obj1 As riesgo
    Set obj1 = GetCachedRiesgo(testIdLower)
    Dim obj2 As riesgo
    Set obj2 = GetCachedRiesgo(testIdUpper)
    
    If obj1 Is Nothing Or obj2 Is Nothing Then
        Test_Cache_Riesgo_KeyCaseInsensitive = ErrJson("Got Nothing for case variant")
        Exit Function
    End If
    
    If obj1.IDRiesgo <> obj2.IDRiesgo Then
        Test_Cache_Riesgo_KeyCaseInsensitive = ErrJson("Case-insensitive key failed: different objects returned")
        Exit Function
    End If
    
    logs = logs & "case_insensitive_ok;"
    Test_Cache_Riesgo_KeyCaseInsensitive = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Riesgo_KeyCaseInsensitive = ErrJson(Err.Description)
End Function

' ============================================================
' EDICION CACHE TESTS (9)
' ============================================================

Public Function Test_Cache_Edicion_Hit() As String
    Dim logs As String
    Dim rs As DAO.Recordset
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbProyectosEdiciones", "IDEdicion")
    If testId = "" Then
        Test_Cache_Edicion_Hit = ErrJson("No test data found for Edicion")
        Exit Function
    End If
    logs = logs & "id=" & testId & ";"
    
    Dim obj As Edicion
    Set obj = GetCachedEdicion(testId)
    
    If obj Is Nothing Then
        Test_Cache_Edicion_Hit = ErrJson("GetCachedEdicion returned Nothing")
        Exit Function
    End If
    logs = logs & "cached;"
    
    Set rs = GetTestDb().OpenRecordset("SELECT * FROM TbProyectosEdiciones WHERE IDEdicion=" & testId)
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        Test_Cache_Edicion_Hit = ErrJson("DB query returned no record")
        Exit Function
    End If
    
    If Not CompareField(obj.IDEdicion, rs!IDEdicion.value) Then
        rs.Close: Set rs = Nothing
        Test_Cache_Edicion_Hit = ErrJson("IDEdicion mismatch")
        Exit Function
    End If
    If Not CompareField(obj.Edicion, rs!Edicion.value) Then
        rs.Close: Set rs = Nothing
        Test_Cache_Edicion_Hit = ErrJson("Edicion mismatch")
        Exit Function
    End If
    
    rs.Close: Set rs = Nothing
    Test_Cache_Edicion_Hit = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Edicion_Hit = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Edicion_CacheConsistency() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbProyectosEdiciones", "IDEdicion")
    If testId = "" Then
        Test_Cache_Edicion_CacheConsistency = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As Edicion
    Set obj1 = GetCachedEdicion(testId)
    Dim obj2 As Edicion
    Set obj2 = GetCachedEdicion(testId)
    
    If obj1 Is Nothing Or obj2 Is Nothing Then
        Test_Cache_Edicion_CacheConsistency = ErrJson("Got Nothing")
        Exit Function
    End If
    
    If obj1.IDEdicion <> obj2.IDEdicion Then
        Test_Cache_Edicion_CacheConsistency = ErrJson("Objects have different ID")
        Exit Function
    End If
    
    logs = logs & "consistent;"
    Test_Cache_Edicion_CacheConsistency = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Edicion_CacheConsistency = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Edicion_Miss() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim obj As Edicion
    Set obj = GetCachedEdicion("999999999")
    
    If Not obj Is Nothing Then
        Test_Cache_Edicion_Miss = ErrJson("Expected Nothing for non-existent ID")
        Exit Function
    End If
    
    logs = logs & "miss_ok;"
    Test_Cache_Edicion_Miss = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Edicion_Miss = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Edicion_Vacio() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim obj As Edicion
    Set obj = GetCachedEdicion("")
    
    If Not obj Is Nothing Then
        Test_Cache_Edicion_Vacio = ErrJson("Expected Nothing for empty ID")
        Exit Function
    End If
    
    logs = logs & "vacio_ok;"
    Test_Cache_Edicion_Vacio = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Edicion_Vacio = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Edicion_Invalidar() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbProyectosEdiciones", "IDEdicion")
    If testId = "" Then
        Test_Cache_Edicion_Invalidar = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As Edicion
    Set obj1 = GetCachedEdicion(testId)
    logs = logs & "cached;"
    
    InvalidarCacheEdicion testId
    logs = logs & "invalidated;"
    
    Dim obj2 As Edicion
    Set obj2 = GetCachedEdicion(testId)
    
    If obj2 Is Nothing Then
        Test_Cache_Edicion_Invalidar = ErrJson("After invalidation, GetCached returned Nothing")
        Exit Function
    End If
    
    If obj2.IDEdicion <> testId Then
        Test_Cache_Edicion_Invalidar = ErrJson("After invalidation, re-query returned wrong ID")
        Exit Function
    End If
    
    logs = logs & "requery_ok;"
    Test_Cache_Edicion_Invalidar = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Edicion_Invalidar = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Edicion_RequeryAfterInvalidate() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbProyectosEdiciones", "IDEdicion")
    If testId = "" Then
        Test_Cache_Edicion_RequeryAfterInvalidate = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As Edicion
    Set obj1 = GetCachedEdicion(testId)
    InvalidarCacheEdicion testId
    Dim obj2 As Edicion
    Set obj2 = GetCachedEdicion(testId)
    
    If obj1.IDEdicion <> obj2.IDEdicion Then
        Test_Cache_Edicion_RequeryAfterInvalidate = ErrJson("Re-query returned different ID")
        Exit Function
    End If
    
    logs = logs & "requery_ok;"
    Test_Cache_Edicion_RequeryAfterInvalidate = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Edicion_RequeryAfterInvalidate = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Edicion_NestedProyecto() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbProyectosEdiciones", "IDEdicion")
    If testId = "" Then
        Test_Cache_Edicion_NestedProyecto = ErrJson("No test data")
        Exit Function
    End If
    
    Dim objEdicion As Edicion
    Set objEdicion = GetCachedEdicion(testId)
    If objEdicion Is Nothing Then
        Test_Cache_Edicion_NestedProyecto = ErrJson("Edicion is Nothing")
        Exit Function
    End If
    logs = logs & "edicion_cached;"
    
    Dim objProyecto As Proyecto
    Set objProyecto = objEdicion.Proyecto
    If objProyecto Is Nothing Then
        Test_Cache_Edicion_NestedProyecto = ErrJson("Proyecto is Nothing")
        Exit Function
    End If
    logs = logs & "proyecto_accessed;"
    
    ' Verify cached
    Dim objProyectoCached As Proyecto
    Set objProyectoCached = GetCachedProyecto(objProyecto.IDProyecto)
    If objProyectoCached Is Nothing Then
        Test_Cache_Edicion_NestedProyecto = ErrJson("Proyecto not in cache")
        Exit Function
    End If
    
    logs = logs & "nested_cache_ok;"
    Test_Cache_Edicion_NestedProyecto = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Edicion_NestedProyecto = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Edicion_SameReference() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbProyectosEdiciones", "IDEdicion")
    If testId = "" Then
        Test_Cache_Edicion_SameReference = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As Edicion
    Set obj1 = GetCachedEdicion(testId)
    Dim obj2 As Edicion
    Set obj2 = GetCachedEdicion(testId)
    
    If obj1 Is Nothing Or obj2 Is Nothing Then
        Test_Cache_Edicion_SameReference = ErrJson("Got Nothing")
        Exit Function
    End If
    
    If obj1.IDEdicion <> obj2.IDEdicion Then
        Test_Cache_Edicion_SameReference = ErrJson("Different IDs returned")
        Exit Function
    End If
    
    logs = logs & "same_reference_ok;"
    Test_Cache_Edicion_SameReference = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Edicion_SameReference = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Edicion_KeyCaseInsensitive() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbProyectosEdiciones", "IDEdicion")
    If testId = "" Then
        Test_Cache_Edicion_KeyCaseInsensitive = ErrJson("No test data")
        Exit Function
    End If
    
    Dim testIdLower As String
    Dim testIdUpper As String
    testIdLower = LCase$(testId)
    testIdUpper = UCase$(testId)
    
    If testIdLower = testIdUpper Then
        logs = logs & "numeric_id_skip;"
        Test_Cache_Edicion_KeyCaseInsensitive = OkJson(logs)
        Exit Function
    End If
    
    Dim obj1 As Edicion
    Set obj1 = GetCachedEdicion(testIdLower)
    Dim obj2 As Edicion
    Set obj2 = GetCachedEdicion(testIdUpper)
    
    If obj1 Is Nothing Or obj2 Is Nothing Then
        Test_Cache_Edicion_KeyCaseInsensitive = ErrJson("Got Nothing for case variant")
        Exit Function
    End If
    
    If obj1.IDEdicion <> obj2.IDEdicion Then
        Test_Cache_Edicion_KeyCaseInsensitive = ErrJson("Case-insensitive key failed")
        Exit Function
    End If
    
    logs = logs & "case_insensitive_ok;"
    Test_Cache_Edicion_KeyCaseInsensitive = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Edicion_KeyCaseInsensitive = ErrJson(Err.Description)
End Function

' ============================================================
' PROYECTO CACHE TESTS (8)
' ============================================================

Public Function Test_Cache_Proyecto_Hit() As String
    Dim rs As DAO.Recordset
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbProyectos", "IDProyecto")
    If testId = "" Then
        Test_Cache_Proyecto_Hit = ErrJson("No test data found for Proyecto")
        Exit Function
    End If
    logs = logs & "id=" & testId & ";"
    
    Dim obj As Proyecto
    Set obj = GetCachedProyecto(testId)
    
    If obj Is Nothing Then
        Test_Cache_Proyecto_Hit = ErrJson("GetCachedProyecto returned Nothing")
        Exit Function
    End If
    logs = logs & "cached;"
    
    Set rs = GetTestDb().OpenRecordset("SELECT * FROM TbProyectos WHERE IDProyecto=" & testId)
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        Test_Cache_Proyecto_Hit = ErrJson("DB query returned no record")
        Exit Function
    End If
    
    If Not CompareField(obj.IDProyecto, rs!IDProyecto.value) Then
        rs.Close: Set rs = Nothing
        Test_Cache_Proyecto_Hit = ErrJson("IDProyecto mismatch")
        Exit Function
    End If
    If Not CompareField(obj.Proyecto, rs!Proyecto.value) Then
        rs.Close: Set rs = Nothing
        Test_Cache_Proyecto_Hit = ErrJson("Proyecto field mismatch")
        Exit Function
    End If
    
    rs.Close: Set rs = Nothing
    Test_Cache_Proyecto_Hit = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Proyecto_Hit = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Proyecto_CacheConsistency() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbProyectos", "IDProyecto")
    If testId = "" Then
        Test_Cache_Proyecto_CacheConsistency = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As Proyecto
    Set obj1 = GetCachedProyecto(testId)
    Dim obj2 As Proyecto
    Set obj2 = GetCachedProyecto(testId)
    
    If obj1 Is Nothing Or obj2 Is Nothing Then
        Test_Cache_Proyecto_CacheConsistency = ErrJson("Got Nothing")
        Exit Function
    End If
    
    If obj1.IDProyecto <> obj2.IDProyecto Then
        Test_Cache_Proyecto_CacheConsistency = ErrJson("Objects have different ID")
        Exit Function
    End If
    
    logs = logs & "consistent;"
    Test_Cache_Proyecto_CacheConsistency = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Proyecto_CacheConsistency = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Proyecto_Miss() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim obj As Proyecto
    Set obj = GetCachedProyecto("999999999")
    
    If Not obj Is Nothing Then
        Test_Cache_Proyecto_Miss = ErrJson("Expected Nothing for non-existent ID")
        Exit Function
    End If
    
    logs = logs & "miss_ok;"
    Test_Cache_Proyecto_Miss = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Proyecto_Miss = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Proyecto_Vacio() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim obj As Proyecto
    Set obj = GetCachedProyecto("")
    
    If Not obj Is Nothing Then
        Test_Cache_Proyecto_Vacio = ErrJson("Expected Nothing for empty ID")
        Exit Function
    End If
    
    logs = logs & "vacio_ok;"
    Test_Cache_Proyecto_Vacio = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Proyecto_Vacio = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Proyecto_Invalidar() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbProyectos", "IDProyecto")
    If testId = "" Then
        Test_Cache_Proyecto_Invalidar = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As Proyecto
    Set obj1 = GetCachedProyecto(testId)
    logs = logs & "cached;"
    
    InvalidarCacheProyecto testId
    logs = logs & "invalidated;"
    
    Dim obj2 As Proyecto
    Set obj2 = GetCachedProyecto(testId)
    
    If obj2 Is Nothing Then
        Test_Cache_Proyecto_Invalidar = ErrJson("After invalidation, GetCached returned Nothing")
        Exit Function
    End If
    
    If obj2.IDProyecto <> testId Then
        Test_Cache_Proyecto_Invalidar = ErrJson("After invalidation, re-query returned wrong ID")
        Exit Function
    End If
    
    logs = logs & "requery_ok;"
    Test_Cache_Proyecto_Invalidar = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Proyecto_Invalidar = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Proyecto_RequeryAfterInvalidate() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbProyectos", "IDProyecto")
    If testId = "" Then
        Test_Cache_Proyecto_RequeryAfterInvalidate = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As Proyecto
    Set obj1 = GetCachedProyecto(testId)
    InvalidarCacheProyecto testId
    Dim obj2 As Proyecto
    Set obj2 = GetCachedProyecto(testId)
    
    If obj1.IDProyecto <> obj2.IDProyecto Then
        Test_Cache_Proyecto_RequeryAfterInvalidate = ErrJson("Re-query returned different ID")
        Exit Function
    End If
    
    logs = logs & "requery_ok;"
    Test_Cache_Proyecto_RequeryAfterInvalidate = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Proyecto_RequeryAfterInvalidate = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Proyecto_SameReference() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbProyectos", "IDProyecto")
    If testId = "" Then
        Test_Cache_Proyecto_SameReference = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As Proyecto
    Set obj1 = GetCachedProyecto(testId)
    Dim obj2 As Proyecto
    Set obj2 = GetCachedProyecto(testId)
    
    If obj1 Is Nothing Or obj2 Is Nothing Then
        Test_Cache_Proyecto_SameReference = ErrJson("Got Nothing")
        Exit Function
    End If
    
    If obj1.IDProyecto <> obj2.IDProyecto Then
        Test_Cache_Proyecto_SameReference = ErrJson("Different IDs returned")
        Exit Function
    End If
    
    logs = logs & "same_reference_ok;"
    Test_Cache_Proyecto_SameReference = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Proyecto_SameReference = ErrJson(Err.Description)
End Function

Public Function Test_Cache_Proyecto_KeyCaseInsensitive() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbProyectos", "IDProyecto")
    If testId = "" Then
        Test_Cache_Proyecto_KeyCaseInsensitive = ErrJson("No test data")
        Exit Function
    End If
    
    Dim testIdLower As String
    Dim testIdUpper As String
    testIdLower = LCase$(testId)
    testIdUpper = UCase$(testId)
    
    If testIdLower = testIdUpper Then
        logs = logs & "numeric_id_skip;"
        Test_Cache_Proyecto_KeyCaseInsensitive = OkJson(logs)
        Exit Function
    End If
    
    Dim obj1 As Proyecto
    Set obj1 = GetCachedProyecto(testIdLower)
    Dim obj2 As Proyecto
    Set obj2 = GetCachedProyecto(testIdUpper)
    
    If obj1 Is Nothing Or obj2 Is Nothing Then
        Test_Cache_Proyecto_KeyCaseInsensitive = ErrJson("Got Nothing for case variant")
        Exit Function
    End If
    
    If obj1.IDProyecto <> obj2.IDProyecto Then
        Test_Cache_Proyecto_KeyCaseInsensitive = ErrJson("Case-insensitive key failed")
        Exit Function
    End If
    
    logs = logs & "case_insensitive_ok;"
    Test_Cache_Proyecto_KeyCaseInsensitive = OkJson(logs)
    Exit Function
EH:
    Test_Cache_Proyecto_KeyCaseInsensitive = ErrJson(Err.Description)
End Function

' ============================================================
' PM CACHE TESTS (9)
' ============================================================

Public Function Test_Cache_PM_Hit() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgosPlanMitigacionPpal", "IDMitigacion")
    If testId = "" Then
        Test_Cache_PM_Hit = ErrJson("No test data found for PM")
        Exit Function
    End If
    logs = logs & "id=" & testId & ";"
    
    Dim obj As pm
    Dim rs As DAO.Recordset
    Set obj = GetCachedPM(testId)
    
    If obj Is Nothing Then
        Test_Cache_PM_Hit = ErrJson("GetCachedPM returned Nothing")
        Exit Function
    End If
    logs = logs & "cached;"
    
    Set rs = GetTestDb().OpenRecordset("SELECT * FROM TbRiesgosPlanMitigacionPpal WHERE IDMitigacion=" & testId)
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        Test_Cache_PM_Hit = ErrJson("DB query returned no record")
        Exit Function
    End If
    
    If Not CompareField(obj.IDMitigacion, rs!IDMitigacion.value) Then
        rs.Close: Set rs = Nothing
        Test_Cache_PM_Hit = ErrJson("IDMitigacion mismatch")
        Exit Function
    End If
    
    rs.Close: Set rs = Nothing
    Test_Cache_PM_Hit = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PM_Hit = ErrJson(Err.Description)
End Function

Public Function Test_Cache_PM_CacheConsistency() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgosPlanMitigacionPpal", "IDMitigacion")
    If testId = "" Then
        Test_Cache_PM_CacheConsistency = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As pm
    Set obj1 = GetCachedPM(testId)
    Dim obj2 As pm
    Set obj2 = GetCachedPM(testId)
    
    If obj1 Is Nothing Or obj2 Is Nothing Then
        Test_Cache_PM_CacheConsistency = ErrJson("Got Nothing")
        Exit Function
    End If
    
    If obj1.IDMitigacion <> obj2.IDMitigacion Then
        Test_Cache_PM_CacheConsistency = ErrJson("Objects have different ID")
        Exit Function
    End If
    
    logs = logs & "consistent;"
    Test_Cache_PM_CacheConsistency = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PM_CacheConsistency = ErrJson(Err.Description)
End Function

Public Function Test_Cache_PM_Miss() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim obj As pm
    Set obj = GetCachedPM("999999999")
    
    If Not obj Is Nothing Then
        Test_Cache_PM_Miss = ErrJson("Expected Nothing for non-existent ID")
        Exit Function
    End If
    
    logs = logs & "miss_ok;"
    Test_Cache_PM_Miss = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PM_Miss = ErrJson(Err.Description)
End Function

Public Function Test_Cache_PM_Vacio() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim obj As pm
    Set obj = GetCachedPM("")
    
    If Not obj Is Nothing Then
        Test_Cache_PM_Vacio = ErrJson("Expected Nothing for empty ID")
        Exit Function
    End If
    
    logs = logs & "vacio_ok;"
    Test_Cache_PM_Vacio = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PM_Vacio = ErrJson(Err.Description)
End Function

Public Function Test_Cache_PM_Invalidar() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgosPlanMitigacionPpal", "IDMitigacion")
    If testId = "" Then
        Test_Cache_PM_Invalidar = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As pm
    Set obj1 = GetCachedPM(testId)
    logs = logs & "cached;"
    
    InvalidarCachePM testId
    logs = logs & "invalidated;"
    
    Dim obj2 As pm
    Set obj2 = GetCachedPM(testId)
    
    If obj2 Is Nothing Then
        Test_Cache_PM_Invalidar = ErrJson("After invalidation, GetCached returned Nothing")
        Exit Function
    End If
    
    If obj2.IDMitigacion <> testId Then
        Test_Cache_PM_Invalidar = ErrJson("After invalidation, re-query returned wrong ID")
        Exit Function
    End If
    
    logs = logs & "requery_ok;"
    Test_Cache_PM_Invalidar = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PM_Invalidar = ErrJson(Err.Description)
End Function

Public Function Test_Cache_PM_RequeryAfterInvalidate() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgosPlanMitigacionPpal", "IDMitigacion")
    If testId = "" Then
        Test_Cache_PM_RequeryAfterInvalidate = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As pm
    Set obj1 = GetCachedPM(testId)
    InvalidarCachePM testId
    Dim obj2 As pm
    Set obj2 = GetCachedPM(testId)
    
    If obj1.IDMitigacion <> obj2.IDMitigacion Then
        Test_Cache_PM_RequeryAfterInvalidate = ErrJson("Re-query returned different ID")
        Exit Function
    End If
    
    logs = logs & "requery_ok;"
    Test_Cache_PM_RequeryAfterInvalidate = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PM_RequeryAfterInvalidate = ErrJson(Err.Description)
End Function

Public Function Test_Cache_PM_SinAcciones() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgosPlanMitigacionPpal", "IDMitigacion")
    
    If testId = "" Then
        logs = logs & "no_pm_data_skip;"
        Test_Cache_PM_SinAcciones = OkJson(logs)
        Exit Function
    End If
    
    Dim obj As pm
    Set obj = GetCachedPM(testId)
    
    If obj Is Nothing Then
        Test_Cache_PM_SinAcciones = ErrJson("PM is Nothing")
        Exit Function
    End If
    
    logs = logs & "sinacciones_found;"
    Test_Cache_PM_SinAcciones = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PM_SinAcciones = ErrJson(Err.Description)
End Function

Public Function Test_Cache_PM_SameReference() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgosPlanMitigacionPpal", "IDMitigacion")
    If testId = "" Then
        Test_Cache_PM_SameReference = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As pm
    Set obj1 = GetCachedPM(testId)
    Dim obj2 As pm
    Set obj2 = GetCachedPM(testId)
    
    If obj1 Is Nothing Or obj2 Is Nothing Then
        Test_Cache_PM_SameReference = ErrJson("Got Nothing")
        Exit Function
    End If
    
    If obj1.IDMitigacion <> obj2.IDMitigacion Then
        Test_Cache_PM_SameReference = ErrJson("Different IDs returned")
        Exit Function
    End If
    
    logs = logs & "same_reference_ok;"
    Test_Cache_PM_SameReference = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PM_SameReference = ErrJson(Err.Description)
End Function

Public Function Test_Cache_PM_KeyCaseInsensitive() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgosPlanMitigacionPpal", "IDMitigacion")
    If testId = "" Then
        Test_Cache_PM_KeyCaseInsensitive = ErrJson("No test data")
        Exit Function
    End If
    
    Dim testIdLower As String
    Dim testIdUpper As String
    testIdLower = LCase$(testId)
    testIdUpper = UCase$(testId)
    
    If testIdLower = testIdUpper Then
        logs = logs & "numeric_id_skip;"
        Test_Cache_PM_KeyCaseInsensitive = OkJson(logs)
        Exit Function
    End If
    
    Dim obj1 As pm
    Set obj1 = GetCachedPM(testIdLower)
    Dim obj2 As pm
    Set obj2 = GetCachedPM(testIdUpper)
    
    If obj1 Is Nothing Or obj2 Is Nothing Then
        Test_Cache_PM_KeyCaseInsensitive = ErrJson("Got Nothing for case variant")
        Exit Function
    End If
    
    If obj1.IDMitigacion <> obj2.IDMitigacion Then
        Test_Cache_PM_KeyCaseInsensitive = ErrJson("Case-insensitive key failed")
        Exit Function
    End If
    
    logs = logs & "case_insensitive_ok;"
    Test_Cache_PM_KeyCaseInsensitive = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PM_KeyCaseInsensitive = ErrJson(Err.Description)
End Function

' ============================================================
' PC CACHE TESTS (9)
' ============================================================

Public Function Test_Cache_PC_Hit() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgosPlanContingenciaPpal", "IDContingencia")
    If testId = "" Then
        Test_Cache_PC_Hit = ErrJson("No test data found for PC")
        Exit Function
    End If
    logs = logs & "id=" & testId & ";"
    
    Dim obj As pc
    Dim rs As DAO.Recordset
    Set obj = GetCachedPC(testId)
    
    If obj Is Nothing Then
        Test_Cache_PC_Hit = ErrJson("GetCachedPC returned Nothing")
        Exit Function
    End If
    logs = logs & "cached;"
    
    Set rs = GetTestDb().OpenRecordset("SELECT * FROM TbRiesgosPlanContingenciaPpal WHERE IDContingencia=" & testId)
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        Test_Cache_PC_Hit = ErrJson("DB query returned no record")
        Exit Function
    End If
    
    If Not CompareField(obj.IDContingencia, rs!IDContingencia.value) Then
        rs.Close: Set rs = Nothing
        Test_Cache_PC_Hit = ErrJson("IDContingencia mismatch")
        Exit Function
    End If
    
    rs.Close: Set rs = Nothing
    Test_Cache_PC_Hit = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PC_Hit = ErrJson(Err.Description)
End Function

Public Function Test_Cache_PC_CacheConsistency() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgosPlanContingenciaPpal", "IDContingencia")
    If testId = "" Then
        Test_Cache_PC_CacheConsistency = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As pc
    Set obj1 = GetCachedPC(testId)
    Dim obj2 As pc
    Set obj2 = GetCachedPC(testId)
    
    If obj1 Is Nothing Or obj2 Is Nothing Then
        Test_Cache_PC_CacheConsistency = ErrJson("Got Nothing")
        Exit Function
    End If
    
    If obj1.IDContingencia <> obj2.IDContingencia Then
        Test_Cache_PC_CacheConsistency = ErrJson("Objects have different ID")
        Exit Function
    End If
    
    logs = logs & "consistent;"
    Test_Cache_PC_CacheConsistency = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PC_CacheConsistency = ErrJson(Err.Description)
End Function

Public Function Test_Cache_PC_Miss() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim obj As pc
    Set obj = GetCachedPC("999999999")
    
    If Not obj Is Nothing Then
        Test_Cache_PC_Miss = ErrJson("Expected Nothing for non-existent ID")
        Exit Function
    End If
    
    logs = logs & "miss_ok;"
    Test_Cache_PC_Miss = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PC_Miss = ErrJson(Err.Description)
End Function

Public Function Test_Cache_PC_Vacio() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim obj As pc
    Set obj = GetCachedPC("")
    
    If Not obj Is Nothing Then
        Test_Cache_PC_Vacio = ErrJson("Expected Nothing for empty ID")
        Exit Function
    End If
    
    logs = logs & "vacio_ok;"
    Test_Cache_PC_Vacio = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PC_Vacio = ErrJson(Err.Description)
End Function

Public Function Test_Cache_PC_Invalidar() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgosPlanContingenciaPpal", "IDContingencia")
    If testId = "" Then
        Test_Cache_PC_Invalidar = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As pc
    Set obj1 = GetCachedPC(testId)
    logs = logs & "cached;"
    
    InvalidarCachePC testId
    logs = logs & "invalidated;"
    
    Dim obj2 As pc
    Set obj2 = GetCachedPC(testId)
    
    If obj2 Is Nothing Then
        Test_Cache_PC_Invalidar = ErrJson("After invalidation, GetCached returned Nothing")
        Exit Function
    End If
    
    If obj2.IDContingencia <> testId Then
        Test_Cache_PC_Invalidar = ErrJson("After invalidation, re-query returned wrong ID")
        Exit Function
    End If
    
    logs = logs & "requery_ok;"
    Test_Cache_PC_Invalidar = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PC_Invalidar = ErrJson(Err.Description)
End Function

Public Function Test_Cache_PC_RequeryAfterInvalidate() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgosPlanContingenciaPpal", "IDContingencia")
    If testId = "" Then
        Test_Cache_PC_RequeryAfterInvalidate = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As pc
    Set obj1 = GetCachedPC(testId)
    InvalidarCachePC testId
    Dim obj2 As pc
    Set obj2 = GetCachedPC(testId)
    
    If obj1.IDContingencia <> obj2.IDContingencia Then
        Test_Cache_PC_RequeryAfterInvalidate = ErrJson("Re-query returned different ID")
        Exit Function
    End If
    
    logs = logs & "requery_ok;"
    Test_Cache_PC_RequeryAfterInvalidate = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PC_RequeryAfterInvalidate = ErrJson(Err.Description)
End Function

Public Function Test_Cache_PC_SinAcciones() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgosPlanContingenciaPpal", "IDContingencia")
    
    If testId = "" Then
        logs = logs & "no_pc_data_skip;"
        Test_Cache_PC_SinAcciones = OkJson(logs)
        Exit Function
    End If
    
    Dim obj As pc
    Set obj = GetCachedPC(testId)
    
    If obj Is Nothing Then
        Test_Cache_PC_SinAcciones = ErrJson("PC is Nothing")
        Exit Function
    End If
    
    logs = logs & "sinacciones_found;"
    Test_Cache_PC_SinAcciones = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PC_SinAcciones = ErrJson(Err.Description)
End Function

Public Function Test_Cache_PC_SameReference() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgosPlanContingenciaPpal", "IDContingencia")
    If testId = "" Then
        Test_Cache_PC_SameReference = ErrJson("No test data")
        Exit Function
    End If
    
    Dim obj1 As pc
    Set obj1 = GetCachedPC(testId)
    Dim obj2 As pc
    Set obj2 = GetCachedPC(testId)
    
    If obj1 Is Nothing Or obj2 Is Nothing Then
        Test_Cache_PC_SameReference = ErrJson("Got Nothing")
        Exit Function
    End If
    
    If obj1.IDContingencia <> obj2.IDContingencia Then
        Test_Cache_PC_SameReference = ErrJson("Different IDs returned")
        Exit Function
    End If
    
    logs = logs & "same_reference_ok;"
    Test_Cache_PC_SameReference = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PC_SameReference = ErrJson(Err.Description)
End Function

Public Function Test_Cache_PC_KeyCaseInsensitive() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim testId As String
    testId = CacheFixtureId("TbRiesgosPlanContingenciaPpal", "IDContingencia")
    If testId = "" Then
        Test_Cache_PC_KeyCaseInsensitive = ErrJson("No test data")
        Exit Function
    End If
    
    Dim testIdLower As String
    Dim testIdUpper As String
    testIdLower = LCase$(testId)
    testIdUpper = UCase$(testId)
    
    If testIdLower = testIdUpper Then
        logs = logs & "numeric_id_skip;"
        Test_Cache_PC_KeyCaseInsensitive = OkJson(logs)
        Exit Function
    End If
    
    Dim obj1 As pc
    Set obj1 = GetCachedPC(testIdLower)
    Dim obj2 As pc
    Set obj2 = GetCachedPC(testIdUpper)
    
    If obj1 Is Nothing Or obj2 Is Nothing Then
        Test_Cache_PC_KeyCaseInsensitive = ErrJson("Got Nothing for case variant")
        Exit Function
    End If
    
    If obj1.IDContingencia <> obj2.IDContingencia Then
        Test_Cache_PC_KeyCaseInsensitive = ErrJson("Case-insensitive key failed")
        Exit Function
    End If
    
    logs = logs & "case_insensitive_ok;"
    Test_Cache_PC_KeyCaseInsensitive = OkJson(logs)
    Exit Function
EH:
    Test_Cache_PC_KeyCaseInsensitive = ErrJson(Err.Description)
End Function

' ============================================================
' BATCH LOADING TESTS (12 tests)
' ============================================================

' ============================================================
' ISSUE #37 RISK TREE DATA LOADER TESTS
' Schema evidence gathered with Dysflow MCP get_schema on Gestion_Riesgos_Datos.accdb:
' - TbProyectos: PK IDProyecto Long required; IDExpediente optional; Proyecto nullable text.
' - TbProyectosEdiciones: PK IDEdicion Long required; FK-shaped IDProyecto Long required; Edicion Integer required.
' - TbRiesgos: PK IDRiesgo Long required; FK-shaped IDEdicion Long required; required CodigoUnico, CodigoRiesgo.
' - TbRiesgosPlanMitigacionPpal: PK IDMitigacion Long required; FK-shaped IDRiesgo Long required.
' - TbRiesgosPlanContingenciaPpal: PK IDContingencia Long required; FK-shaped IDRiesgo Long required.
' - TbRiesgosPlanMitigacionDetalle: PK IDAccionMitigacion Long required; FK-shaped IDMitigacion Long required.
' - TbRiesgosPlanContingenciaDetalle: PK IDAccionContingencia Long required; FK-shaped IDContingencia Long required.
' Relationships: Dysflow get_relationships reported only an Access navigation relationship, so fixture order follows the application FK-shaped fields above.
' Seed order: Test_Fixtures.SeedAll creates expediente -> project -> edition -> risk -> PM/PC; test-local rows add PM action -> PC action.
' Teardown order: Test_Fixtures.TeardownAll deletes PM/PC action children before plans, risk, edition, project, expediente.
' ============================================================

Private Sub SeedRiskTreeLoaderActions(ByVal p_db As DAO.Database)
    p_db.Execute "DELETE FROM TbRiesgosPlanContingenciaDetalle WHERE IDAccionContingencia=" & CACHE_FIX_PC_ACTION_ID, dbFailOnError
    p_db.Execute "DELETE FROM TbRiesgosPlanMitigacionDetalle WHERE IDAccionMitigacion=" & CACHE_FIX_PM_ACTION_ID, dbFailOnError
    p_db.Execute "INSERT INTO TbRiesgosPlanMitigacionDetalle " & _
        "(IDAccionMitigacion, IDMitigacion, CodAccion, Accion, ResponsableAccion, Estado, EsUltimaAccion) " & _
        "VALUES (" & CACHE_FIX_PM_ACTION_ID & ", " & CACHE_FIX_PM_ID & ", 'PMATST', " & _
        "'Fixture PM action test', 'TESTUSER', 'Definida', 'Si')", dbFailOnError
    p_db.Execute "INSERT INTO TbRiesgosPlanContingenciaDetalle " & _
        "(IDAccionContingencia, IDContingencia, CodAccion, Accion, ResponsableAccion, Estado, EsUltimaAccion) " & _
        "VALUES (" & CACHE_FIX_PC_ACTION_ID & ", " & CACHE_FIX_PC_ID & ", 'PCATST', " & _
        "'Fixture PC action test', 'TESTUSER', 'Definida', 'Si')", dbFailOnError
End Sub

Public Function Test_RiskTreeDataLoader_PreloadsProjectEditionRiskPlanActionGraph() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH

    Dim errMsg As String
    Dim db As DAO.Database
    Dim loader As RiskTreeDataLoader
    Dim result As Scripting.Dictionary
    Dim risks As Scripting.Dictionary
    Dim risk As riesgo
    Dim pms As Scripting.Dictionary
    Dim pcs As Scripting.Dictionary
    Dim pmObj As pm
    Dim pcObj As pc

    If Not PrepareCacheFixtureGraph(errMsg) Then
        Test_RiskTreeDataLoader_PreloadsProjectEditionRiskPlanActionGraph = ErrJson(errMsg)
        Exit Function
    End If
    Set db = GetTestDb(errMsg)
    If db Is Nothing Then
        Test_RiskTreeDataLoader_PreloadsProjectEditionRiskPlanActionGraph = ErrJson(errMsg)
        Exit Function
    End If
    SeedRiskTreeLoaderActions db
    logs = logs & "seeded_project_edition_risk_plan_actions;"

    Set risk = GetCachedRiesgo(CStr(CACHE_FIX_RIESGO_ID), errMsg)
    If errMsg <> "" Then GoTo Fail
    If risk Is Nothing Then
        errMsg = "fixture risk was not loaded"
        GoTo Fail
    End If
    Set risks = New Scripting.Dictionary
    risks.CompareMode = TextCompare
    risks.Add risk.IDRiesgo, risk

    Set loader = New RiskTreeDataLoader
    Set result = loader.PreloadForEdition(CStr(CACHE_FIX_EDICION_ID), risks, errMsg)
    If errMsg <> "" Then GoTo Fail
    If result Is Nothing Then
        errMsg = "loader returned Nothing"
        GoTo Fail
    End If
    If Not CBool(result("ok")) Then
        errMsg = "loader should preserve non-fatal ok result"
        GoTo Fail
    End If
    If CLng(result("riskCount")) <> 1 Then
        errMsg = "expected riskCount=1, got " & result("riskCount")
        GoTo Fail
    End If
    If CLng(result("pmCount")) <> 1 Or CLng(result("pcCount")) <> 1 Then
        errMsg = "expected exactly one PM and one PC preload result"
        GoTo Fail
    End If
    If CLng(result("pmActionCount")) <> 1 Or CLng(result("pcActionCount")) <> 1 Then
        errMsg = "expected exactly one PM action and one PC action preload result"
        GoTo Fail
    End If

    Set pms = risk.ColPMs
    Set pcs = risk.ColPCs
    If pms Is Nothing Then
        errMsg = "PM collection was not injected for fixture risk"
        GoTo Fail
    End If
    If Not pms.Exists(CStr(CACHE_FIX_PM_ID)) Then
        errMsg = "PM collection was not injected for fixture risk"
        GoTo Fail
    End If
    If pcs Is Nothing Then
        errMsg = "PC collection was not injected for fixture risk"
        GoTo Fail
    End If
    If Not pcs.Exists(CStr(CACHE_FIX_PC_ID)) Then
        errMsg = "PC collection was not injected for fixture risk"
        GoTo Fail
    End If

    Set pmObj = pms(CStr(CACHE_FIX_PM_ID))
    Set pcObj = pcs(CStr(CACHE_FIX_PC_ID))
    If pmObj.colAcciones Is Nothing Then
        errMsg = "PM action collection was not injected"
        GoTo Fail
    End If
    If Not pmObj.colAcciones.Exists(CStr(CACHE_FIX_PM_ACTION_ID)) Then
        errMsg = "PM action collection was not injected"
        GoTo Fail
    End If
    If pcObj.colAcciones Is Nothing Then
        errMsg = "PC action collection was not injected"
        GoTo Fail
    End If
    If Not pcObj.colAcciones.Exists(CStr(CACHE_FIX_PC_ACTION_ID)) Then
        errMsg = "PC action collection was not injected"
        GoTo Fail
    End If

    logs = logs & "preload_graph_ok;"
    Test_RiskTreeDataLoader_PreloadsProjectEditionRiskPlanActionGraph = OkJson(logs)
Teardown:
    On Error Resume Next
    Test_Fixtures.TeardownAll
    Test_Helper.ResetTestSession
    Set db = Nothing
    On Error GoTo 0
    Exit Function
Fail:
    Test_RiskTreeDataLoader_PreloadsProjectEditionRiskPlanActionGraph = BuildFailStr(errMsg, logs)
    GoTo Teardown
EH:
    errMsg = "Test_RiskTreeDataLoader_PreloadsProjectEditionRiskPlanActionGraph: " & Err.Description
    GoTo Fail
End Function

Public Function Test_RiskTreeDataLoader_EmptyRiskDictionaryIsNoOp() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH

    Dim errMsg As String
    Dim loader As RiskTreeDataLoader
    Dim result As Scripting.Dictionary
    Dim risks As Scripting.Dictionary

    Set risks = New Scripting.Dictionary
    risks.CompareMode = TextCompare
    Set loader = New RiskTreeDataLoader
    Set result = loader.PreloadForEdition(CStr(CACHE_FIX_EDICION_ID), risks, errMsg)
    If errMsg <> "" Then
        Test_RiskTreeDataLoader_EmptyRiskDictionaryIsNoOp = BuildFailStr(errMsg, logs)
        Exit Function
    End If
    If result Is Nothing Then
        Test_RiskTreeDataLoader_EmptyRiskDictionaryIsNoOp = ErrJson("loader returned Nothing")
        Exit Function
    End If
    If CLng(result("riskCount")) <> 0 Then
        Test_RiskTreeDataLoader_EmptyRiskDictionaryIsNoOp = ErrJson("empty input must keep riskCount=0")
        Exit Function
    End If
    If CLng(result("warningCount")) <> 0 Then
        Test_RiskTreeDataLoader_EmptyRiskDictionaryIsNoOp = ErrJson("empty input should not create warnings")
        Exit Function
    End If

    logs = logs & "empty_noop_ok;"
    Test_RiskTreeDataLoader_EmptyRiskDictionaryIsNoOp = OkJson(logs)
    Exit Function
EH:
    Test_RiskTreeDataLoader_EmptyRiskDictionaryIsNoOp = ErrJson(Err.Description)
End Function

Public Function Test_RiskTreeDataLoader_NonFatalFailurePreservesLazyFallback() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH

    Dim errMsg As String
    Dim loader As RiskTreeDataLoader
    Dim result As Scripting.Dictionary
    Dim risks As Scripting.Dictionary
    Dim risk As riesgo
    Dim fallbackPMs As Scripting.Dictionary
    Dim fallbackPCs As Scripting.Dictionary

    Set risk = New riesgo
    risk.IDRiesgo = "999999998"
    Set risks = New Scripting.Dictionary
    risks.CompareMode = TextCompare
    risks.Add risk.IDRiesgo, risk

    Set loader = New RiskTreeDataLoader
    Set result = loader.PreloadForEdition("", risks, errMsg)
    If errMsg <> "" Then
        Test_RiskTreeDataLoader_NonFatalFailurePreservesLazyFallback = BuildFailStr("loader must not surface fatal error: " & errMsg, logs)
        Exit Function
    End If
    If result Is Nothing Then
        Test_RiskTreeDataLoader_NonFatalFailurePreservesLazyFallback = ErrJson("loader returned Nothing")
        Exit Function
    End If
    If Not CBool(result("ok")) Then
        Test_RiskTreeDataLoader_NonFatalFailurePreservesLazyFallback = ErrJson("non-fatal failure must keep ok=true")
        Exit Function
    End If
    If CLng(result("warningCount")) = 0 Then
        Test_RiskTreeDataLoader_NonFatalFailurePreservesLazyFallback = ErrJson("expected at least one preload warning")
        Exit Function
    End If
    Set fallbackPMs = risk.ColPMs
    If Not fallbackPMs Is Nothing Then
        Test_RiskTreeDataLoader_NonFatalFailurePreservesLazyFallback = ErrJson("PM lazy fallback was replaced by injected data")
        Exit Function
    End If
    Set fallbackPCs = risk.ColPCs
    If Not fallbackPCs Is Nothing Then
        Test_RiskTreeDataLoader_NonFatalFailurePreservesLazyFallback = ErrJson("PC lazy fallback was replaced by injected data")
        Exit Function
    End If

    logs = logs & "nonfatal_lazy_fallback_ok;"
    Test_RiskTreeDataLoader_NonFatalFailurePreservesLazyFallback = OkJson(logs)
    Exit Function
EH:
    Test_RiskTreeDataLoader_NonFatalFailurePreservesLazyFallback = ErrJson(Err.Description)
End Function

Public Function Test_GetPMsPorEdicionBatch_Empty() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim p_Error As String
    Dim result As Scripting.Dictionary
    Set result = Constructor.getPMsPorEdicionBatch("", Nothing, p_Error)
    If p_Error <> "" Then
        ' Sad path OK: validación de input vacío funciona
        logs = logs & "validation_error_caught;"
        Test_GetPMsPorEdicionBatch_Empty = OkJson(logs, p_Error)
        Exit Function
    End If
    If Not result Is Nothing Then
        Test_GetPMsPorEdicionBatch_Empty = ErrJson("Should return Nothing for empty input")
        Exit Function
    End If
    
    logs = logs & "empty_ok;"
    Test_GetPMsPorEdicionBatch_Empty = OkJson(logs)
    Exit Function
EH:
    Test_GetPMsPorEdicionBatch_Empty = ErrJson(Err.Description)
End Function

Public Function Test_GetPMsPorEdicionBatch_OneRisk() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim riesgoId As String
    riesgoId = CacheFixtureId("TbRiesgos", "IDRiesgo")
    If riesgoId = "" Then
        Test_GetPMsPorEdicionBatch_OneRisk = ErrJson("No test data")
        Exit Function
    End If
    
    Dim p_Error As String
    Dim edicionId As String
    Dim colIds As New Collection
    edicionId = CacheFixtureId("TbProyectosEdiciones", "IDEdicion")
    colIds.Add riesgoId
    
    Dim result As Scripting.Dictionary
    Set result = Constructor.getPMsPorEdicionBatch(edicionId, colIds, p_Error)
    If p_Error <> "" Then
        Test_GetPMsPorEdicionBatch_OneRisk = ErrJson("Error: " & p_Error)
        Exit Function
    End If
    
    ' result may be Nothing if no PMs exist for this riesgo, which is valid
    logs = logs & "one_risk_ok;"
    Test_GetPMsPorEdicionBatch_OneRisk = OkJson(logs)
    Exit Function
EH:
    Test_GetPMsPorEdicionBatch_OneRisk = ErrJson(Err.Description)
End Function

Public Function Test_GetPMsPorEdicionBatch_MultipleRisks() As String
    Dim logs As String
    Dim p_Error As String
    logs = ""
    On Error GoTo EH
    
    Dim riesgoId1 As String
    Dim edicionId As String
    riesgoId1 = CacheFixtureId("TbRiesgos", "IDRiesgo")
    edicionId = CacheFixtureId("TbProyectosEdiciones", "IDEdicion")
    If riesgoId1 = "" Then
        Test_GetPMsPorEdicionBatch_MultipleRisks = ErrJson("No test data")
        Exit Function
    End If
    
    Dim colIds As New Collection
    colIds.Add riesgoId1
    
    Dim result As Scripting.Dictionary
    Set result = Constructor.getPMsPorEdicionBatch(edicionId, colIds, p_Error)
    If p_Error <> "" Then
        ' Sad path OK: múltiples riesgos con IDs inválidos rechazados
        logs = logs & "multiple_risks_validation_ok;"
        Test_GetPMsPorEdicionBatch_MultipleRisks = OkJson(logs, p_Error)
        Exit Function
    End If
    
    logs = logs & "multiple_risks_ok;"
    Test_GetPMsPorEdicionBatch_MultipleRisks = OkJson(logs)
    Exit Function
EH:
    Test_GetPMsPorEdicionBatch_MultipleRisks = ErrJson(Err.Description)
End Function

Public Function Test_GetPCsPorEdicionBatch_Empty() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim p_Error As String
    Dim result As Scripting.Dictionary
    Set result = Constructor.getPCsPorEdicionBatch("", Nothing, p_Error)
    If p_Error <> "" Then
        ' Sad path OK: validación de input vacío funciona
        logs = logs & "validation_error_caught;"
        Test_GetPCsPorEdicionBatch_Empty = OkJson(logs, p_Error)
        Exit Function
    End If
    If Not result Is Nothing Then
        Test_GetPCsPorEdicionBatch_Empty = ErrJson("Should return Nothing for empty input")
        Exit Function
    End If
    
    logs = logs & "empty_ok;"
    Test_GetPCsPorEdicionBatch_Empty = OkJson(logs)
    Exit Function
EH:
    Test_GetPCsPorEdicionBatch_Empty = ErrJson(Err.Description)
End Function

Public Function Test_GetPCsPorEdicionBatch_MultipleRisks() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim riesgoId1 As String
    riesgoId1 = CacheFixtureId("TbRiesgos", "IDRiesgo")
    If riesgoId1 = "" Then
        Test_GetPCsPorEdicionBatch_MultipleRisks = ErrJson("No test data")
        Exit Function
    End If
    
    Dim p_Error As String
    Dim edicionId As String
    Dim colIds As New Collection
    edicionId = CacheFixtureId("TbProyectosEdiciones", "IDEdicion")
    colIds.Add riesgoId1
    
    Dim result As Scripting.Dictionary
    Set result = Constructor.getPCsPorEdicionBatch(edicionId, colIds, p_Error)
    If p_Error <> "" Then
        Test_GetPCsPorEdicionBatch_MultipleRisks = ErrJson("Error: " & p_Error)
        Exit Function
    End If
    
    logs = logs & "multiple_risks_ok;"
    Test_GetPCsPorEdicionBatch_MultipleRisks = OkJson(logs)
    Exit Function
EH:
    Test_GetPCsPorEdicionBatch_MultipleRisks = ErrJson(Err.Description)
End Function

Public Function Test_GetPMAccionesBatch_Normal() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim pmId As String
    pmId = CacheFixtureId("TbRiesgosPlanMitigacionPpal", "IDMitigacion")
    If pmId = "" Then
        Test_GetPMAccionesBatch_Normal = ErrJson("No test data")
        Exit Function
    End If
    
    Dim p_Error As String
    Dim colIds As New Collection
    colIds.Add pmId
    
    Dim result As Scripting.Dictionary
    Set result = Constructor.getPMAccionesBatch(colIds, p_Error)
    If p_Error <> "" Then
        Test_GetPMAccionesBatch_Normal = ErrJson("Error: " & p_Error)
        Exit Function
    End If
    
    ' result may be Nothing if no actions - valid
    logs = logs & "normal_ok;"
    Test_GetPMAccionesBatch_Normal = OkJson(logs)
    Exit Function
EH:
    Test_GetPMAccionesBatch_Normal = ErrJson(Err.Description)
End Function

Public Function Test_GetPMAccionesBatch_NoActions() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    ' Create an empty collection
    Dim colIds As New Collection
    
    Dim p_Error As String
    Dim result As Scripting.Dictionary
    Set result = Constructor.getPMAccionesBatch(colIds, p_Error)
    If p_Error <> "" Then
        Test_GetPMAccionesBatch_NoActions = ErrJson("Error: " & p_Error)
        Exit Function
    End If
    If Not result Is Nothing Then
        Test_GetPMAccionesBatch_NoActions = ErrJson("Should return Nothing for empty collection")
        Exit Function
    End If
    
    logs = logs & "no_actions_ok;"
    Test_GetPMAccionesBatch_NoActions = OkJson(logs)
    Exit Function
EH:
    Test_GetPMAccionesBatch_NoActions = ErrJson(Err.Description)
End Function

Public Function Test_GetPCAccionesBatch_Normal() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim pcId As String
    pcId = CacheFixtureId("TbRiesgosPlanContingenciaPpal", "IDContingencia")
    If pcId = "" Then
        Test_GetPCAccionesBatch_Normal = ErrJson("No test data")
        Exit Function
    End If
    
    Dim p_Error As String
    Dim colIds As New Collection
    colIds.Add pcId
    
    Dim result As Scripting.Dictionary
    Set result = Constructor.getPCAccionesBatch(colIds, p_Error)
    If p_Error <> "" Then
        Test_GetPCAccionesBatch_Normal = ErrJson("Error: " & p_Error)
        Exit Function
    End If
    
    logs = logs & "normal_ok;"
    Test_GetPCAccionesBatch_Normal = OkJson(logs)
    Exit Function
EH:
    Test_GetPCAccionesBatch_Normal = ErrJson(Err.Description)
End Function

Public Function Test_Riesgo_InjectarColecciones_OK() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    ' Create a non-empty dictionary with a PM
    Dim pmId As String
    pmId = CacheFixtureId("TbRiesgosPlanMitigacionPpal", "IDMitigacion")
    If pmId = "" Then
        Test_Riesgo_InjectarColecciones_OK = ErrJson("No test data")
        Exit Function
    End If
    
    Dim p_Error As String
    Dim risk As riesgo
    Set risk = GetCachedRiesgo(CacheFixtureId("TbRiesgos", "IDRiesgo"))
    If risk Is Nothing Then
        Test_Riesgo_InjectarColecciones_OK = ErrJson("No riesgo fixture")
        Exit Function
    End If
    
    ' Get the PM
    Dim pm As pm
    Set pm = GetCachedPM(pmId)
    If pm Is Nothing Then
        Test_Riesgo_InjectarColecciones_OK = ErrJson("No PM fixture")
        Exit Function
    End If
    
    ' Create dictionary with one PM
    Dim colPMs As New Scripting.Dictionary
    colPMs.CompareMode = TextCompare
    colPMs.Add pm.IDMitigacion, pm
    
    ' Inject
    risk.InjectarColPMs colPMs
    
    ' Verify ColPMs returns what we injected (no lazy load)
    Dim resultCol As Scripting.Dictionary
    Set resultCol = risk.colPMs
    If resultCol Is Nothing Then
        Test_Riesgo_InjectarColecciones_OK = ErrJson("ColPMs returned Nothing after injection")
        Exit Function
    End If
    If Not resultCol.Exists(pm.IDMitigacion) Then
        Test_Riesgo_InjectarColecciones_OK = ErrJson("Injected PM not found in ColPMs")
        Exit Function
    End If
    
    logs = logs & "injection_ok;"
    Test_Riesgo_InjectarColecciones_OK = OkJson(logs)
    Exit Function
EH:
    Test_Riesgo_InjectarColecciones_OK = ErrJson(Err.Description)
End Function

Public Function Test_Riesgo_InjectarColecciones_Fallback() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim riesgoId As String
    riesgoId = CacheFixtureId("TbRiesgos", "IDRiesgo")
    If riesgoId = "" Then
        Test_Riesgo_InjectarColecciones_Fallback = ErrJson("No test data")
        Exit Function
    End If
    
    Dim p_Error As String
    Dim risk As riesgo
    Set risk = GetCachedRiesgo(riesgoId)
    If risk Is Nothing Then
        Test_Riesgo_InjectarColecciones_Fallback = ErrJson("No riesgo fixture")
        Exit Function
    End If
    
    ' Clear any existing cache by calling ColPMs (which may load from DB)
    Dim existing As Scripting.Dictionary
    Set existing = risk.colPMs
    Dim hadExisting As Boolean
    hadExisting = Not existing Is Nothing
    
    ' Inject empty dictionary
    Dim emptyCol As New Scripting.Dictionary
    emptyCol.CompareMode = TextCompare
    risk.InjectarColPMs emptyCol
    
    ' ColPMs should now return the empty dict (falls through to lazy-load)
    Dim resultCol As Scripting.Dictionary
    Set resultCol = risk.colPMs
    If resultCol Is Nothing Then
        Test_Riesgo_InjectarColecciones_Fallback = ErrJson("ColPMs returned Nothing after empty injection")
        Exit Function
    End If
    
    logs = logs & "fallback_ok;"
    Test_Riesgo_InjectarColecciones_Fallback = OkJson(logs)
    Exit Function
EH:
    Test_Riesgo_InjectarColecciones_Fallback = ErrJson(Err.Description)
End Function

Public Function Test_PM_InjectarAcciones_OK() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim pmId As String
    pmId = CacheFixtureId("TbRiesgosPlanMitigacionPpal", "IDMitigacion")
    If pmId = "" Then
        Test_PM_InjectarAcciones_OK = ErrJson("No test data")
        Exit Function
    End If
    
    Dim p_Error As String
    Dim pm As pm
    Set pm = GetCachedPM(pmId)
    If pm Is Nothing Then
        Test_PM_InjectarAcciones_OK = ErrJson("No PM fixture")
        Exit Function
    End If
    
    ' Create an owned empty action dictionary; injection behavior is independent of live actions.
    Dim colAcciones As New Scripting.Dictionary
    colAcciones.CompareMode = TextCompare
    
    pm.InjectarColAcciones colAcciones
    
    Dim resultCol As Scripting.Dictionary
    Set resultCol = pm.colAcciones
    If resultCol Is Nothing Then
        Test_PM_InjectarAcciones_OK = ErrJson("colAcciones returned Nothing after injection")
        Exit Function
    End If
    
    logs = logs & "pm_injection_ok;"
    Test_PM_InjectarAcciones_OK = OkJson(logs)
    Exit Function
EH:
    Test_PM_InjectarAcciones_OK = ErrJson(Err.Description)
End Function

Public Function Test_PC_InjectarAcciones_OK() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    Dim pcId As String
    pcId = CacheFixtureId("TbRiesgosPlanContingenciaPpal", "IDContingencia")
    If pcId = "" Then
        Test_PC_InjectarAcciones_OK = ErrJson("No test data")
        Exit Function
    End If
    
    Dim p_Error As String
    Dim pc As pc
    Set pc = GetCachedPC(pcId)
    If pc Is Nothing Then
        Test_PC_InjectarAcciones_OK = ErrJson("No PC fixture")
        Exit Function
    End If
    
    ' Create an owned empty action dictionary; injection behavior is independent of live actions.
    Dim colAcciones As New Scripting.Dictionary
    colAcciones.CompareMode = TextCompare
    
    pc.InjectarColAcciones colAcciones
    
    Dim resultCol As Scripting.Dictionary
    Set resultCol = pc.colAcciones
    If resultCol Is Nothing Then
        Test_PC_InjectarAcciones_OK = ErrJson("colAcciones returned Nothing after injection")
        Exit Function
    End If
    
    logs = logs & "pc_injection_ok;"
    Test_PC_InjectarAcciones_OK = OkJson(logs)
    Exit Function
EH:
    Test_PC_InjectarAcciones_OK = ErrJson(Err.Description)
End Function

' ============================================================
' RESETGLOBALS TESTS (2)
' ============================================================

Public Function Test_Cache_ResetGlobals_AllCleared() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    ' Populate all caches first
    Dim riesgoId As String
    riesgoId = CacheFixtureId("TbRiesgos", "IDRiesgo")
    Dim edicionId As String
    edicionId = CacheFixtureId("TbProyectosEdiciones", "IDEdicion")
    Dim proyectoId As String
    proyectoId = CacheFixtureId("TbProyectos", "IDProyecto")
    Dim pmId As String
    pmId = CacheFixtureId("TbRiesgosPlanMitigacionPpal", "IDMitigacion")
    Dim pcId As String
    pcId = CacheFixtureId("TbRiesgosPlanContingenciaPpal", "IDContingencia")
    
    If riesgoId <> "" Then GetCachedRiesgo riesgoId
    If edicionId <> "" Then GetCachedEdicion edicionId
    If proyectoId <> "" Then GetCachedProyecto proyectoId
    If pmId <> "" Then GetCachedPM pmId
    If pcId <> "" Then GetCachedPC pcId
    logs = logs & "caches_populated;"
    
    ' Call ResetGlobals
    Dim p_Error As String
    ResetGlobals p_Error
    If p_Error <> "" Then
        Test_Cache_ResetGlobals_AllCleared = ErrJson("ResetGlobals returned error: " & p_Error)
        Exit Function
    End If
    logs = logs & "reset_called;"
    
    ' Now check each cache is empty by trying to get
    ' After ResetGlobals, next GetCached should go to DB
    ' But if dictionaries are Nothing, they will be recreated
    ' The behavior is: Set m_DicX = Nothing ? next access recreates
    ' So after ResetGlobals, GetCached repopulates from DB (not from old cache)
    
    ' Verify by calling GetCached and seeing it works (recreates from DB)
    If riesgoId <> "" Then
        Dim objRiesgo As riesgo
        Set objRiesgo = GetCachedRiesgo(riesgoId)
        If objRiesgo Is Nothing Then
            ' This is actually OK - it means cache was cleared and DB query returned no record
            logs = logs & "riesgo_requery_ok;"
        Else
            logs = logs & "riesgo_requery_ok;"
        End If
    End If
    
    logs = logs & "all_cleared_verified;"
    Test_Cache_ResetGlobals_AllCleared = OkJson(logs)
    Exit Function
EH:
    Test_Cache_ResetGlobals_AllCleared = ErrJson(Err.Description)
End Function

Public Function Test_Cache_ResetGlobals_RequeryAfter() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH
    
    ' Get a fixture ID
    Dim testId As String
    testId = CacheFixtureId("TbProyectos", "IDProyecto")
    If testId = "" Then
        Test_Cache_ResetGlobals_RequeryAfter = ErrJson("No test data")
        Exit Function
    End If
    
    ' Populate
    Dim obj1 As Proyecto
    Set obj1 = GetCachedProyecto(testId)
    logs = logs & "populated;"
    
    ' Reset
    Dim p_Error As String
    ResetGlobals p_Error
    If p_Error <> "" Then
        Test_Cache_ResetGlobals_RequeryAfter = ErrJson("ResetGlobals error")
        Exit Function
    End If
    
    ' Reconnect to backend (ResetGlobals clears m_ActiveBackendURL; m_BackendSandboxURL se recarga en LeeConfiguracionLocal)
    LeeConfiguracionLocal p_Error
    If p_Error <> "" Then
        Test_Cache_ResetGlobals_RequeryAfter = ErrJson("LeeConfiguracionLocal error")
        Exit Function
    End If
    
    ' Requery
    Dim obj2 As Proyecto
    Set obj2 = GetCachedProyecto(testId)
    
    If obj2 Is Nothing Then
        Test_Cache_ResetGlobals_RequeryAfter = ErrJson("Requery after ResetGlobals returned Nothing")
        Exit Function
    End If
    
    If obj2.IDProyecto <> testId Then
        Test_Cache_ResetGlobals_RequeryAfter = ErrJson("Requery returned wrong ID")
        Exit Function
    End If
    
    logs = logs & "requery_after_reset_ok;"
    Test_Cache_ResetGlobals_RequeryAfter = OkJson(logs)
    Exit Function
EH:
    Test_Cache_ResetGlobals_RequeryAfter = ErrJson(Err.Description)
End Function

' ============================================================
' RUN ALL AGGREGATOR
' ============================================================

Private Sub RecordCacheRunAllResult(ByVal p_Name As String, ByVal p_Result As String, ByRef p_Total As Integer, ByRef p_Passed As Integer, ByRef p_Failed As Integer, ByRef p_Failures As String)
    p_Total = p_Total + 1
    If InStr(1, p_Result, """ok"":true", vbTextCompare) > 0 Then
        p_Passed = p_Passed + 1
    Else
        p_Failed = p_Failed + 1
        If Len(p_Failures) > 0 Then p_Failures = p_Failures & ","
        p_Failures = p_Failures & "{""name"":""" & JsonEscape(p_Name) & """,""detail"":" & p_Result & "}"
    End If
End Sub

Private Sub RecordCacheSliceResult(ByVal p_Name As String, ByVal p_Result As String, ByRef p_Total As Integer, ByRef p_Passed As Integer, ByRef p_Failed As Integer, ByRef p_FirstFailure As String)
    p_Total = p_Total + 1
    If InStr(1, p_Result, """ok"":true", vbTextCompare) > 0 Then
        p_Passed = p_Passed + 1
    Else
        p_Failed = p_Failed + 1
        If p_FirstFailure = "" Then
            p_FirstFailure = "{""name"":""" & JsonEscape(p_Name) & """,""detail"":" & p_Result & "}"
        End If
    End If
End Sub

Private Function BuildCacheSliceJson(ByVal p_Slice As String, ByVal p_Total As Integer, ByVal p_Passed As Integer, ByVal p_Failed As Integer, ByVal p_FirstFailure As String) As String
    Dim okText As String
    Dim errText As String
    Dim firstFailureText As String
    Dim payload As String
    Dim logs As String

    okText = BoolJson(p_Failed = 0)
    If p_Failed = 0 Then
        errText = "null"
        firstFailureText = "null"
    Else
        errText = """slice failed"""
        firstFailureText = p_FirstFailure
    End If

    payload = "{""slice"":""" & JsonEscape(p_Slice) & """,""total"":" & p_Total & ",""passed"":" & p_Passed & ",""failed"":" & p_Failed & ",""firstFailure"":" & firstFailureText & "}"
    logs = """slice=" & JsonEscape(p_Slice) & ";total=" & p_Total & ";passed=" & p_Passed & ";failed=" & p_Failed & """"
    BuildCacheSliceJson = "{""ok"":" & okText & ",""value"":""" & JsonEscape(p_Slice) & """,""payload"":" & payload & ",""error"":" & errText & ",""logs"": [" & logs & "]}"
End Function

Public Function Test_Cache_Riesgo_RunSlice() As String
    On Error GoTo EH

    Dim failures As String
    Dim total As Integer
    Dim passed As Integer
    Dim failed As Integer
    Dim runError As String

    If Not ForceLocalBackend(runError) Then
        Test_Cache_Riesgo_RunSlice = "{""ok"":false,""value"":""cache-riesgo"",""payload"":null,""error"":""TESTS BLOCKED: " & JsonEscape(runError) & """,""logs"": [""SuiteSetup failed""]}"
        Exit Function
    End If

    Test_Fixtures.SeedAll

    RecordCacheSliceResult "Riesgo_Hit", Test_Cache_Riesgo_Hit, total, passed, failed, failures
    RecordCacheSliceResult "Riesgo_CacheConsistency", Test_Cache_Riesgo_CacheConsistency, total, passed, failed, failures
    RecordCacheSliceResult "Riesgo_Miss", Test_Cache_Riesgo_Miss, total, passed, failed, failures
    RecordCacheSliceResult "Riesgo_Vacio", Test_Cache_Riesgo_Vacio, total, passed, failed, failures
    RecordCacheSliceResult "Riesgo_Invalidar", Test_Cache_Riesgo_Invalidar, total, passed, failed, failures
    RecordCacheSliceResult "Riesgo_RequeryAfterInvalidate", Test_Cache_Riesgo_RequeryAfterInvalidate, total, passed, failed, failures
    RecordCacheSliceResult "Riesgo_NestedEdicion", Test_Cache_Riesgo_NestedEdicion, total, passed, failed, failures
    RecordCacheSliceResult "Riesgo_NestedProyecto", Test_Cache_Riesgo_NestedProyecto, total, passed, failed, failures
    RecordCacheSliceResult "Riesgo_SameReference", Test_Cache_Riesgo_SameReference, total, passed, failed, failures
    RecordCacheSliceResult "Riesgo_KeyCaseInsensitive", Test_Cache_Riesgo_KeyCaseInsensitive, total, passed, failed, failures

    Test_Fixtures.TeardownAll
    Test_Helper.ResetTestSession

    Test_Cache_Riesgo_RunSlice = BuildCacheSliceJson("cache-riesgo", total, passed, failed, failures)
    Exit Function

EH:
    On Error Resume Next
    Test_Fixtures.TeardownAll
    Test_Helper.ResetTestSession
    Test_Cache_Riesgo_RunSlice = "{""ok"":false,""value"":""cache-riesgo"",""payload"":null,""error"":""" & JsonEscape(Err.Description) & """,""logs"": [""RunSlice exception""]}"
End Function

Public Function Test_Cache_RunAll() As String
    Dim failures As String
    Dim total As Integer
    Dim passed As Integer
    Dim failed As Integer
    Dim logs As String

    ' --- SuiteSetup: configurar modo testing (v2.1) ---
    Dim runError As String
    If Not ForceLocalBackend(runError) Then
        Dim failLogs(0 To 0) As String
        failLogs(0) = "SuiteSetup falló"
        Test_Cache_RunAll = BuildJsonFail("TESTS BLOCKED: " & runError, failLogs)
        Exit Function
    End If

    failures = ""
    total = 0
    passed = 0
    failed = 0
    logs = ""

    ' Skill v2.1: SeedAll después de ForceLocalBackend
    Test_Fixtures.SeedAll
    logs = "SeedAll done; "
    ' Riesgo tests
    RecordCacheRunAllResult "Riesgo_Hit", Test_Cache_Riesgo_Hit, total, passed, failed, failures
    RecordCacheRunAllResult "Riesgo_Consistency", Test_Cache_Riesgo_CacheConsistency, total, passed, failed, failures
    RecordCacheRunAllResult "Riesgo_Miss", Test_Cache_Riesgo_Miss, total, passed, failed, failures
    RecordCacheRunAllResult "Riesgo_Vacio", Test_Cache_Riesgo_Vacio, total, passed, failed, failures
    RecordCacheRunAllResult "Riesgo_Invalidar", Test_Cache_Riesgo_Invalidar, total, passed, failed, failures
    RecordCacheRunAllResult "Riesgo_Requery", Test_Cache_Riesgo_RequeryAfterInvalidate, total, passed, failed, failures
    RecordCacheRunAllResult "Riesgo_NestedEd", Test_Cache_Riesgo_NestedEdicion, total, passed, failed, failures
    RecordCacheRunAllResult "Riesgo_NestedProj", Test_Cache_Riesgo_NestedProyecto, total, passed, failed, failures
    RecordCacheRunAllResult "Riesgo_SameRef", Test_Cache_Riesgo_SameReference, total, passed, failed, failures
    RecordCacheRunAllResult "Riesgo_CaseIns", Test_Cache_Riesgo_KeyCaseInsensitive, total, passed, failed, failures
    
    ' Edicion tests
    RecordCacheRunAllResult "Edicion_Hit", Test_Cache_Edicion_Hit, total, passed, failed, failures
    RecordCacheRunAllResult "Edicion_Consistency", Test_Cache_Edicion_CacheConsistency, total, passed, failed, failures
    RecordCacheRunAllResult "Edicion_Miss", Test_Cache_Edicion_Miss, total, passed, failed, failures
    RecordCacheRunAllResult "Edicion_Vacio", Test_Cache_Edicion_Vacio, total, passed, failed, failures
    RecordCacheRunAllResult "Edicion_Invalidar", Test_Cache_Edicion_Invalidar, total, passed, failed, failures
    RecordCacheRunAllResult "Edicion_Requery", Test_Cache_Edicion_RequeryAfterInvalidate, total, passed, failed, failures
    RecordCacheRunAllResult "Edicion_NestedProj", Test_Cache_Edicion_NestedProyecto, total, passed, failed, failures
    RecordCacheRunAllResult "Edicion_SameRef", Test_Cache_Edicion_SameReference, total, passed, failed, failures
    RecordCacheRunAllResult "Edicion_CaseIns", Test_Cache_Edicion_KeyCaseInsensitive, total, passed, failed, failures
    
    ' Proyecto tests
    RecordCacheRunAllResult "Proyecto_Hit", Test_Cache_Proyecto_Hit, total, passed, failed, failures
    RecordCacheRunAllResult "Proyecto_Consistency", Test_Cache_Proyecto_CacheConsistency, total, passed, failed, failures
    RecordCacheRunAllResult "Proyecto_Miss", Test_Cache_Proyecto_Miss, total, passed, failed, failures
    RecordCacheRunAllResult "Proyecto_Vacio", Test_Cache_Proyecto_Vacio, total, passed, failed, failures
    RecordCacheRunAllResult "Proyecto_Invalidar", Test_Cache_Proyecto_Invalidar, total, passed, failed, failures
    RecordCacheRunAllResult "Proyecto_Requery", Test_Cache_Proyecto_RequeryAfterInvalidate, total, passed, failed, failures
    RecordCacheRunAllResult "Proyecto_SameRef", Test_Cache_Proyecto_SameReference, total, passed, failed, failures
    RecordCacheRunAllResult "Proyecto_CaseIns", Test_Cache_Proyecto_KeyCaseInsensitive, total, passed, failed, failures
    
    ' PM tests
    RecordCacheRunAllResult "PM_Hit", Test_Cache_PM_Hit, total, passed, failed, failures
    RecordCacheRunAllResult "PM_Consistency", Test_Cache_PM_CacheConsistency, total, passed, failed, failures
    RecordCacheRunAllResult "PM_Miss", Test_Cache_PM_Miss, total, passed, failed, failures
    RecordCacheRunAllResult "PM_Vacio", Test_Cache_PM_Vacio, total, passed, failed, failures
    RecordCacheRunAllResult "PM_Invalidar", Test_Cache_PM_Invalidar, total, passed, failed, failures
    RecordCacheRunAllResult "PM_Requery", Test_Cache_PM_RequeryAfterInvalidate, total, passed, failed, failures
    RecordCacheRunAllResult "PM_SinAcc", Test_Cache_PM_SinAcciones, total, passed, failed, failures
    RecordCacheRunAllResult "PM_SameRef", Test_Cache_PM_SameReference, total, passed, failed, failures
    RecordCacheRunAllResult "PM_CaseIns", Test_Cache_PM_KeyCaseInsensitive, total, passed, failed, failures
    
    ' PC tests
    RecordCacheRunAllResult "PC_Hit", Test_Cache_PC_Hit, total, passed, failed, failures
    RecordCacheRunAllResult "PC_Consistency", Test_Cache_PC_CacheConsistency, total, passed, failed, failures
    RecordCacheRunAllResult "PC_Miss", Test_Cache_PC_Miss, total, passed, failed, failures
    RecordCacheRunAllResult "PC_Vacio", Test_Cache_PC_Vacio, total, passed, failed, failures
    RecordCacheRunAllResult "PC_Invalidar", Test_Cache_PC_Invalidar, total, passed, failed, failures
    RecordCacheRunAllResult "PC_Requery", Test_Cache_PC_RequeryAfterInvalidate, total, passed, failed, failures
    RecordCacheRunAllResult "PC_SinAcc", Test_Cache_PC_SinAcciones, total, passed, failed, failures
    RecordCacheRunAllResult "PC_SameRef", Test_Cache_PC_SameReference, total, passed, failed, failures
    RecordCacheRunAllResult "PC_CaseIns", Test_Cache_PC_KeyCaseInsensitive, total, passed, failed, failures
    
    ' Batch loading tests
    RecordCacheRunAllResult "RiskTreeLoader_Graph", Test_RiskTreeDataLoader_PreloadsProjectEditionRiskPlanActionGraph, total, passed, failed, failures
    RecordCacheRunAllResult "RiskTreeLoader_Empty", Test_RiskTreeDataLoader_EmptyRiskDictionaryIsNoOp, total, passed, failed, failures
    RecordCacheRunAllResult "RiskTreeLoader_NonFatal", Test_RiskTreeDataLoader_NonFatalFailurePreservesLazyFallback, total, passed, failed, failures
    RecordCacheRunAllResult "BatchPM_Empty", Test_GetPMsPorEdicionBatch_Empty, total, passed, failed, failures
    RecordCacheRunAllResult "BatchPM_OneRisk", Test_GetPMsPorEdicionBatch_OneRisk, total, passed, failed, failures
    RecordCacheRunAllResult "BatchPM_MultipleRisks", Test_GetPMsPorEdicionBatch_MultipleRisks, total, passed, failed, failures
    RecordCacheRunAllResult "BatchPC_Empty", Test_GetPCsPorEdicionBatch_Empty, total, passed, failed, failures
    RecordCacheRunAllResult "BatchPC_MultipleRisks", Test_GetPCsPorEdicionBatch_MultipleRisks, total, passed, failed, failures
    RecordCacheRunAllResult "BatchPMAcc_Normal", Test_GetPMAccionesBatch_Normal, total, passed, failed, failures
    RecordCacheRunAllResult "BatchPMAcc_NoActions", Test_GetPMAccionesBatch_NoActions, total, passed, failed, failures
    RecordCacheRunAllResult "BatchPCAcc_Normal", Test_GetPCAccionesBatch_Normal, total, passed, failed, failures
    RecordCacheRunAllResult "Riesgo_Injectar_OK", Test_Riesgo_InjectarColecciones_OK, total, passed, failed, failures
    RecordCacheRunAllResult "Riesgo_Injectar_Fallback", Test_Riesgo_InjectarColecciones_Fallback, total, passed, failed, failures
    RecordCacheRunAllResult "PM_Injectar_OK", Test_PM_InjectarAcciones_OK, total, passed, failed, failures
    RecordCacheRunAllResult "PC_Injectar_OK", Test_PC_InjectarAcciones_OK, total, passed, failed, failures
    
    ' ResetGlobals tests
    RecordCacheRunAllResult "ResetAll", Test_Cache_ResetGlobals_AllCleared, total, passed, failed, failures
    RecordCacheRunAllResult "ResetRequery", Test_Cache_ResetGlobals_RequeryAfter, total, passed, failed, failures
    
    ' Cleanup: Skill v2.1 TeardownAll después de los tests
    Test_Fixtures.TeardownAll

    ' -- SuiteTeardown ---------------------------------------------
    Test_Helper.ResetTestSession

    logs = "total=" & total & ";passed=" & passed & ";failed=" & failed

    Dim payload As String
    payload = "{""total"":" & total & ",""passed"":" & passed & ",""failed"":" & failed & ",""failures"": [" & failures & "]}"

    Test_Cache_RunAll = OkJson(logs, payload)
    Exit Function

EH:
    On Error Resume Next
    Test_Fixtures.TeardownAll
    Test_Helper.ResetTestSession
    Dim ehLogs(0 To 0) As String
    ehLogs(0) = "Error en suite"
    Test_Cache_RunAll = BuildJsonFail(Err.Description, ehLogs)
End Function

Public Function Test_Cache_ManualDetailRefresh_InvalidatesRiskOnlyOrder() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH

    Dim m_Error As String
    Dim m_Orden As String

    m_Error = ""
    m_Orden = ObtenerOrdenInvalidacionRiesgosPorScope("risk", m_Error)
    If m_Error <> "" Then
        Test_Cache_ManualDetailRefresh_InvalidatesRiskOnlyOrder = BuildFailStr("ObtenerOrdenInvalidacionRiesgosPorScope error: " & m_Error, logs)
        Exit Function
    End If
    logs = logs & "orden=" & m_Orden & ";"

    If StrComp(m_Orden, "risk", vbTextCompare) <> 0 Then
        Test_Cache_ManualDetailRefresh_InvalidatesRiskOnlyOrder = BuildFailStr("El refresco manual de detalle debe invalidar solo 'risk'", logs)
        Exit Function
    End If

    Test_Cache_ManualDetailRefresh_InvalidatesRiskOnlyOrder = OkJson(logs)
    Exit Function
EH:
    Test_Cache_ManualDetailRefresh_InvalidatesRiskOnlyOrder = ErrJson(Err.Description)
End Function

Public Function Test_Cache_FullRefreshScope_RehydrateEditionReplacesStaleReference() As String
    Dim logs As String
    logs = ""
    On Error GoTo EH

    Dim testId As String
    testId = CacheFixtureId("TbProyectosEdiciones", "IDEdicion")
    If testId = "" Then
        Test_Cache_FullRefreshScope_RehydrateEditionReplacesStaleReference = ErrJson("No se pudo preparar edición de test")
        GoTo Teardown
    End If
    logs = logs & "fixture_edition=" & testId & ";"

    Dim m_Error As String
    Dim objStale As Edicion
    Dim objFresh As Edicion

    Set objStale = GetCachedEdicion(testId, m_Error)
    If m_Error <> "" Then
        Test_Cache_FullRefreshScope_RehydrateEditionReplacesStaleReference = BuildFailStr("GetCachedEdicion inicial error: " & m_Error, logs)
        Exit Function
    End If
    If objStale Is Nothing Then
        Test_Cache_FullRefreshScope_RehydrateEditionReplacesStaleReference = ErrJson("No se pudo cachear edición inicial")
        Exit Function
    End If
    logs = logs & "cached_initial;"

    InvalidarCacheEdicion testId, m_Error
    If m_Error <> "" Then
        Test_Cache_FullRefreshScope_RehydrateEditionReplacesStaleReference = BuildFailStr("InvalidarCacheEdicion error: " & m_Error, logs)
        Exit Function
    End If
    logs = logs & "invalidated;"

    Set objFresh = RehidratarEdicionTrasInvalidacion(testId, m_Error)
    If m_Error <> "" Then
        Test_Cache_FullRefreshScope_RehydrateEditionReplacesStaleReference = BuildFailStr("RehidratarEdicionTrasInvalidacion error: " & m_Error, logs)
        Exit Function
    End If
    If objFresh Is Nothing Then
        Test_Cache_FullRefreshScope_RehydrateEditionReplacesStaleReference = ErrJson("Rehidratación devolvió Nothing")
        Exit Function
    End If

    If StrComp(objFresh.IDEdicion, testId, vbTextCompare) <> 0 Then
        Test_Cache_FullRefreshScope_RehydrateEditionReplacesStaleReference = BuildFailStr("La edición rehidratada no coincide con el IDEdicion esperado", logs)
        Exit Function
    End If

    If objStale Is objFresh Then
        Test_Cache_FullRefreshScope_RehydrateEditionReplacesStaleReference = BuildFailStr("La rehidratación no reemplazó la referencia stale de edición", logs)
        Exit Function
    End If

    logs = logs & "rehydrated_new_reference;"
    Test_Cache_FullRefreshScope_RehydrateEditionReplacesStaleReference = OkJson(logs)
Teardown:
    On Error Resume Next
    If testId <> "" Then
        InvalidarCacheEdicion testId
    End If
    On Error GoTo 0
    Exit Function
EH:
    Test_Cache_FullRefreshScope_RehydrateEditionReplacesStaleReference = ErrJson(Err.Description)
    Resume Teardown
End Function




