Attribute VB_Name = "Test_BackendCache"
Option Compare Database
Option Explicit

' =============================================================================
' Test module: Test_BackendCache
' Phase 3 of SDD change staging-alignment-prueba-001 (Issue #22)
'
' Covers axis (iii) of Phase 3 — stale cached connection invalidation in
' getdb() and friends. Scenarios exercised:
'   1. CloseCachedBackendConnection clears the cache (Test 1)
'   2. m_ActiveBackendURL changes at runtime → getdb() must reopen (Test 2 — the canary)
'   3. LeeConfiguracionLocal re-reads config → cache must be dropped (Test 3)
'   4. m_TestingMode toggled without ResetTestSession → next getdb() returns
'      the configured non-testing backend, not stale sandbox (Test 4)
'   5. AssertSandboxBackend must reject when the cached connection points
'      to a different sandbox URL than the active one (Test 5)
'
' All tests follow access-vba-tdd v2.4.2:
'   - Public Function returning canonical JSON (BuildJsonOk / BuildJsonFail)
'   - Zero arguments
'   - Explicit logs() with Arrange/Act/Assert steps
'   - No Debug.Print, no MsgBox, no UI
'   - Sandbox safety via temp .accdb files (CreateIsolatedTempDb pattern)
'   - TestOnlySetBackendConfigOverride / TestOnlyConfigureTestingBackend for setup
'   - TestOnlyOverrideActiveBackendURL is a new test-only helper for runtime URL change
'
' The temp .accdb helpers below are private copies of the canonical
' TestHelper.CreateIsolatedTempDb / DisposeIsolatedTempDb from
' access-vba-tdd v2.4.2 §5.5. They are kept local to this module to honor
' the no-modify scope rule for TestHelper.bas.
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
    MakeIsolatedTempDbPath = Environ$("TEMP") & "\expedientes_test_backendcache_" & p_Suffix & "_" & Format$(Now, "yyyymmddhhnnssfff") & ".accdb"
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

' ---- Test 1: manual close invalidates the cache ----

' Scenario under test:
'   1. getdb() opens a connection and caches it in m_dbCached
'   2. CloseCachedBackendConnection is called explicitly
'   3. The next getdb() must reopen, NOT return the now-closed cached object
'
' Schema/code evidence: Variables Globales.bas lines 791-846 (getdb) and 673-677
' (CloseCachedBackendConnection). BeginTestSession pre-checks are bypassed
' because we set up an isolated temp backend via TestOnlySetBackendConfigOverride
' rather than touching the production sandbox.
Public Function Test_BackendCache_StaleCacheInvalidatedOnManualClose() As String
    Dim logs(0 To 7) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_BackendCache_StaleCacheInvalidatedOnManualClose = JsonFail("test did not complete", logs)

    On Error GoTo HandleError
    ResetTestSession

    Dim errMsg As String
    Dim dbErr As String
    Dim prodPath As String
    Dim dbTemp As DAO.Database
    Dim db1 As DAO.Database
    Dim db2 As DAO.Database

    ' --- Arrange ---
    logs(0) = "1. Arrange: create isolated temp .accdb for the manual-close scenario"
    prodPath = MakeIsolatedTempDbPath("manualclose")
    Set dbTemp = CreateIsolatedTempDb(prodPath, errMsg)
    If dbTemp Is Nothing Then
        Test_BackendCache_StaleCacheInvalidatedOnManualClose = JsonFail("CreateIsolatedTempDb: " & errMsg, logs)
        GoTo Teardown
    End If
    Set dbTemp = Nothing  ' close so getdb() can open it

    logs(1) = "2. Arrange: override config so m_ActiveBackendURL resolves to " & prodPath
    TestOnlyResetBackendConfigOverride
    TestOnlySetBackendConfigOverride "PROD", prodPath, prodPath, prodPath, ""
    LeeConfiguracionLocal errMsg
    If errMsg <> "" Then
        Test_BackendCache_StaleCacheInvalidatedOnManualClose = JsonFail("LeeConfiguracionLocal: " & errMsg, logs)
        GoTo Teardown
    End If

    ' --- Act ---
    logs(2) = "3. Act: first getdb() — should open and cache the connection in m_dbCached"
    Set db1 = getdb(dbErr)
    If db1 Is Nothing Then
        Test_BackendCache_StaleCacheInvalidatedOnManualClose = JsonFail("getdb() #1 returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    logs(3) = "4. Act: CloseCachedBackendConnection — must clear m_dbCached"
    CloseCachedBackendConnection

    logs(4) = "5. Act: second getdb() — must reopen, not return stale/closed"
    Set db2 = getdb(dbErr)
    If db2 Is Nothing Then
        Test_BackendCache_StaleCacheInvalidatedOnManualClose = JsonFail("getdb() #2 returned Nothing after close: " & dbErr, logs)
        GoTo Teardown
    End If

    ' --- Assert ---
    logs(5) = "6. Assert: db1 and db2 are distinct object instances (proves cache was invalidated)"
    If db1 Is db2 Then
        Test_BackendCache_StaleCacheInvalidatedOnManualClose = JsonFail("After close, getdb() returned the SAME object — cache was NOT invalidated", logs)
        GoTo Teardown
    End If

    logs(6) = "7. Assert: db2.Name matches expected prodPath"
    If InStr(1, db2.Name, prodPath, vbTextCompare) = 0 Then
        Test_BackendCache_StaleCacheInvalidatedOnManualClose = JsonFail("db2.Name does not match expected path: " & db2.Name, logs)
        GoTo Teardown
    End If

    Test_BackendCache_StaleCacheInvalidatedOnManualClose = JsonOk("manual_close_invalidates_cache", logs)

Teardown:
    On Error Resume Next
    Set db1 = Nothing
    Set db2 = Nothing
    TestOnlyResetBackendConfigOverride
    CloseCachedBackendConnection
    DisposeIsolatedTempDb prodPath, dbTemp
    Exit Function

HandleError:
    Test_BackendCache_StaleCacheInvalidatedOnManualClose = JsonFail(Err.Description, logs)
    Resume Teardown
End Function

' ---- Test 2: URL change at runtime — the canary ----

' Scenario under test (the core stale-cache bug from explore §3.D):
'   1. Open backend A via getdb() — caches connection
'   2. Change m_ActiveBackendURL to backend B (via TestOnlyOverrideActiveBackendURL)
'   3. Call getdb() again — must detect URL change and reopen against B,
'      NOT return the stale A connection
'
' On the current code, getdb() validates the cache via .Name only (which
' passes if the DAO object is alive) and returns the stale connection —
' this test is RED on the current implementation and goes GREEN after
' the URL-change detection in C3 is applied.
'
' Schema/code evidence: Variables Globales.bas lines 791-846 (getdb) and
' lines 214-225 (m_ActiveBackendURL private). The test uses the new
' TestOnlyOverrideActiveBackendURL helper from C4 to mutate the private
' URL from the test module without going through LeeConfiguracionLocal.
Public Function Test_BackendCache_RejectsStaleCacheWhenBackendURLChanges() As String
    Dim logs(0 To 9) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_BackendCache_RejectsStaleCacheWhenBackendURLChanges = JsonFail("test did not complete", logs)

    On Error GoTo HandleError
    ResetTestSession

    Dim errMsg As String
    Dim dbErr As String
    Dim prodPath As String
    Dim sbxPath As String
    Dim dbTemp1 As DAO.Database
    Dim dbTemp2 As DAO.Database
    Dim db1 As DAO.Database
    Dim db2 As DAO.Database

    ' --- Arrange ---
    logs(0) = "1. Arrange: create two isolated temp .accdb backends (PROD-like and SANDBOX-like)"
    prodPath = MakeIsolatedTempDbPath("stale_prod")
    Set dbTemp1 = CreateIsolatedTempDb(prodPath, errMsg)
    If dbTemp1 Is Nothing Then
        Test_BackendCache_RejectsStaleCacheWhenBackendURLChanges = JsonFail("create prod temp: " & errMsg, logs)
        GoTo Teardown
    End If
    Set dbTemp1 = Nothing

    sbxPath = MakeIsolatedTempDbPath("stale_sbx")
    Set dbTemp2 = CreateIsolatedTempDb(sbxPath, errMsg)
    If dbTemp2 Is Nothing Then
        Test_BackendCache_RejectsStaleCacheWhenBackendURLChanges = JsonFail("create sbx temp: " & errMsg, logs)
        GoTo Teardown
    End If
    Set dbTemp2 = Nothing

    logs(1) = "2. Arrange: override config so initial m_ActiveBackendURL = " & prodPath & " (via PROD key)"
    TestOnlyResetBackendConfigOverride
    TestOnlySetBackendConfigOverride "PROD", prodPath, sbxPath, sbxPath, ""
    LeeConfiguracionLocal errMsg
    If errMsg <> "" Then
        Test_BackendCache_RejectsStaleCacheWhenBackendURLChanges = JsonFail("LeeConfiguracionLocal: " & errMsg, logs)
        GoTo Teardown
    End If

    ' --- Act 1: open initial backend (PROD) ---
    logs(2) = "3. Act: getdb() #1 — opens prodPath, caches connection in m_dbCached"
    Set db1 = getdb(dbErr)
    If db1 Is Nothing Then
        Test_BackendCache_RejectsStaleCacheWhenBackendURLChanges = JsonFail("getdb() #1 failed: " & dbErr, logs)
        GoTo Teardown
    End If
    logs(3) = "4. Assert: db1.Name contains " & prodPath
    If InStr(1, db1.Name, prodPath, vbTextCompare) = 0 Then
        Test_BackendCache_RejectsStaleCacheWhenBackendURLChanges = JsonFail("db1 does not point to prod temp: " & db1.Name, logs)
        GoTo Teardown
    End If

    ' --- Act 2: runtime URL change ---
    logs(4) = "5. Act: TestOnlyOverrideActiveBackendURL → m_ActiveBackendURL = " & sbxPath
    TestOnlyOverrideActiveBackendURL sbxPath
    logs(5) = "5a. Note: GetActiveBackendKey() after override = '" & GetActiveBackendKey() & "' (key is independent of URL)"

    ' --- Act 3: second getdb() — should detect URL change and reopen ---
    logs(6) = "6. Act: getdb() #2 — must detect URL change and reopen against sbxPath. RED on current code: returns stale prodPath."
    Set db2 = getdb(dbErr)
    If db2 Is Nothing Then
        Test_BackendCache_RejectsStaleCacheWhenBackendURLChanges = JsonFail("getdb() #2 returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' --- Assert ---
    logs(7) = "7. Assert: db2.Name must NOT contain prodPath (stale cache must be rejected)"
    If InStr(1, db2.Name, prodPath, vbTextCompare) > 0 Then
        Test_BackendCache_RejectsStaleCacheWhenBackendURLChanges = JsonFail("STALE CACHE BUG: db2 still points to " & prodPath & " after URL change to " & sbxPath, logs)
        GoTo Teardown
    End If

    logs(8) = "8. Assert: db2.Name must contain sbxPath (the new URL)"
    If InStr(1, db2.Name, sbxPath, vbTextCompare) = 0 Then
        Test_BackendCache_RejectsStaleCacheWhenBackendURLChanges = JsonFail("db2.Name does not match " & sbxPath & ": " & db2.Name, logs)
        GoTo Teardown
    End If

    Test_BackendCache_RejectsStaleCacheWhenBackendURLChanges = JsonOk("stale_cache_rejected_on_url_change", logs)

Teardown:
    On Error Resume Next
    Set db1 = Nothing
    Set db2 = Nothing
    TestOnlyResetBackendConfigOverride
    CloseCachedBackendConnection
    DisposeIsolatedTempDb prodPath, dbTemp1
    DisposeIsolatedTempDb sbxPath, dbTemp2
    Exit Function

HandleError:
    Test_BackendCache_RejectsStaleCacheWhenBackendURLChanges = JsonFail(Err.Description, logs)
    Resume Teardown
End Function

' ---- Test 3: LeeConfiguracionLocal must invalidate the cache ----

' Scenario under test (explore §3.C):
'   1. Open backend A via getdb() — caches connection
'   2. Re-override config to point at backend B, then call LeeConfiguracionLocal
'   3. Next getdb() must return B, not stale A
'
' On the current code, LeeConfiguracionLocal does NOT clear m_dbCached
' unless ResetGlobals was called first (only EVE() calls ResetGlobals
' before LeeConfiguracionLocal). This test is RED on the current
' implementation and goes GREEN after the InvalidateBackendCache call
' wired into LeeConfiguracionLocal (C2).
'
' Schema/code evidence: Variables Globales.bas lines 531-626
' (LeeConfiguracionLocal). The fix adds InvalidateBackendCache early in
' the function so the next getdb() re-opens against the freshly-resolved
' URL even when ResetGlobals was not called.
Public Function Test_BackendCache_LeeConfiguracionLocalInvalidatesCache() As String
    Dim logs(0 To 9) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_BackendCache_LeeConfiguracionLocalInvalidatesCache = JsonFail("test did not complete", logs)

    On Error GoTo HandleError
    ResetTestSession

    Dim errMsg As String
    Dim dbErr As String
    Dim path1 As String
    Dim path2 As String
    Dim dbTemp1 As DAO.Database
    Dim dbTemp2 As DAO.Database
    Dim db1 As DAO.Database
    Dim db2 As DAO.Database

    ' --- Arrange ---
    logs(0) = "1. Arrange: create two isolated temp .accdb backends"
    path1 = MakeIsolatedTempDbPath("lee_first")
    Set dbTemp1 = CreateIsolatedTempDb(path1, errMsg)
    If dbTemp1 Is Nothing Then
        Test_BackendCache_LeeConfiguracionLocalInvalidatesCache = JsonFail("create temp1: " & errMsg, logs)
        GoTo Teardown
    End If
    Set dbTemp1 = Nothing

    path2 = MakeIsolatedTempDbPath("lee_second")
    Set dbTemp2 = CreateIsolatedTempDb(path2, errMsg)
    If dbTemp2 Is Nothing Then
        Test_BackendCache_LeeConfiguracionLocalInvalidatesCache = JsonFail("create temp2: " & errMsg, logs)
        GoTo Teardown
    End If
    Set dbTemp2 = Nothing

    ' --- Phase 1: open path1 ---
    logs(1) = "2. Arrange: override config with path1, call LeeConfiguracionLocal"
    TestOnlyResetBackendConfigOverride
    TestOnlySetBackendConfigOverride "PROD", path1, path2, path2, ""
    LeeConfiguracionLocal errMsg
    If errMsg <> "" Then
        Test_BackendCache_LeeConfiguracionLocalInvalidatesCache = JsonFail("LeeConfiguracionLocal #1: " & errMsg, logs)
        GoTo Teardown
    End If

    logs(2) = "3. Act: getdb() #1 — opens path1, caches connection"
    Set db1 = getdb(dbErr)
    If db1 Is Nothing Then
        Test_BackendCache_LeeConfiguracionLocalInvalidatesCache = JsonFail("getdb() #1 failed: " & dbErr, logs)
        GoTo Teardown
    End If
    If InStr(1, db1.Name, path1, vbTextCompare) = 0 Then
        Test_BackendCache_LeeConfiguracionLocalInvalidatesCache = JsonFail("db1 not path1: " & db1.Name, logs)
        GoTo Teardown
    End If

    ' --- Phase 2: re-config to path2 ---
    logs(3) = "4. Act: re-override config with path2, call LeeConfiguracionLocal again (C2 must drop m_dbCached)"
    TestOnlySetBackendConfigOverride "PROD", path2, path1, path1, ""
    LeeConfiguracionLocal errMsg
    If errMsg <> "" Then
        Test_BackendCache_LeeConfiguracionLocalInvalidatesCache = JsonFail("LeeConfiguracionLocal #2: " & errMsg, logs)
        GoTo Teardown
    End If
    logs(4) = "4a. Note: GetActiveBackendKey() after re-config = '" & GetActiveBackendKey() & "'"

    ' --- Act 3: second getdb() — should NOT return stale path1 ---
    logs(5) = "5. Act: getdb() #2 — should open path2, not return cached path1 connection. RED on current code."
    Set db2 = getdb(dbErr)
    If db2 Is Nothing Then
        Test_BackendCache_LeeConfiguracionLocalInvalidatesCache = JsonFail("getdb() #2 failed: " & dbErr, logs)
        GoTo Teardown
    End If

    ' --- Assert ---
    logs(6) = "6. Assert: db2.Name matches path2, not path1"
    If InStr(1, db2.Name, path1, vbTextCompare) > 0 Then
        Test_BackendCache_LeeConfiguracionLocalInvalidatesCache = JsonFail("STALE CACHE BUG: db2 still points to path1 after re-config to path2", logs)
        GoTo Teardown
    End If
    If InStr(1, db2.Name, path2, vbTextCompare) = 0 Then
        Test_BackendCache_LeeConfiguracionLocalInvalidatesCache = JsonFail("db2.Name does not match path2: " & db2.Name, logs)
        GoTo Teardown
    End If

    Test_BackendCache_LeeConfiguracionLocalInvalidatesCache = JsonOk("lee_config_local_invalidates_cache", logs)

Teardown:
    On Error Resume Next
    Set db1 = Nothing
    Set db2 = Nothing
    TestOnlyResetBackendConfigOverride
    CloseCachedBackendConnection
    DisposeIsolatedTempDb path1, dbTemp1
    DisposeIsolatedTempDb path2, dbTemp2
    Exit Function

HandleError:
    Test_BackendCache_LeeConfiguracionLocalInvalidatesCache = JsonFail(Err.Description, logs)
    Resume Teardown
End Function

' ---- Test 4: testing mode toggle must re-resolve the configured backend ----

' Scenario under test (explore §3.B):
'   1. m_ActiveBackendURL = PROD path; m_TestingBackendURL = sandbox path; m_TestingMode = True
'   2. getdb() returns sandbox connection (cached)
'   3. Caller toggles m_TestingMode = False directly (no ResetTestSession)
'   4. Next getdb() must return PROD, NOT the stale sandbox connection
'
' On the current code, m_dbCached holds the sandbox connection (still a
' valid DAO object) and the .Name check passes — getdb() returns the
' sandbox silently. This test is RED on the current implementation
' and goes GREEN after the URL-change detection in C3 is applied.
'
' Schema/code evidence: Variables Globales.bas lines 791-846 (getdb) and
' the m_TestingMode / m_TestingBackendURL / m_ActiveBackendURL globals.
Public Function Test_BackendCache_TestingModeToggleFromTrueToFalseReturnsCorrectBackend() As String
    Dim logs(0 To 9) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_BackendCache_TestingModeToggleFromTrueToFalseReturnsCorrectBackend = JsonFail("test did not complete", logs)

    On Error GoTo HandleError
    ResetTestSession

    Dim errMsg As String
    Dim dbErr As String
    Dim prodPath As String
    Dim sbxPath As String
    Dim dbTemp1 As DAO.Database
    Dim dbTemp2 As DAO.Database
    Dim db1 As DAO.Database
    Dim db2 As DAO.Database

    ' --- Arrange ---
    logs(0) = "1. Arrange: create prod-like and sandbox-like temp .accdb backends"
    prodPath = MakeIsolatedTempDbPath("tog_prod")
    Set dbTemp1 = CreateIsolatedTempDb(prodPath, errMsg)
    If dbTemp1 Is Nothing Then
        Test_BackendCache_TestingModeToggleFromTrueToFalseReturnsCorrectBackend = JsonFail("create prod: " & errMsg, logs)
        GoTo Teardown
    End If
    Set dbTemp1 = Nothing

    sbxPath = MakeIsolatedTempDbPath("tog_sbx")
    Set dbTemp2 = CreateIsolatedTempDb(sbxPath, errMsg)
    If dbTemp2 Is Nothing Then
        Test_BackendCache_TestingModeToggleFromTrueToFalseReturnsCorrectBackend = JsonFail("create sbx: " & errMsg, logs)
        GoTo Teardown
    End If
    Set dbTemp2 = Nothing

    ' --- Setup: m_ActiveBackendURL = prodPath, m_TestingBackendURL = sbxPath ---
    logs(1) = "2. Arrange: override config (BackendActivo=PROD) so m_ActiveBackendURL = " & prodPath
    TestOnlyResetBackendConfigOverride
    TestOnlySetBackendConfigOverride "PROD", prodPath, sbxPath, sbxPath, ""
    LeeConfiguracionLocal errMsg
    If errMsg <> "" Then
        Test_BackendCache_TestingModeToggleFromTrueToFalseReturnsCorrectBackend = JsonFail("LeeConfiguracionLocal: " & errMsg, logs)
        GoTo Teardown
    End If

    logs(2) = "3. Arrange: configure testing backend to point at " & sbxPath & ", enable m_TestingMode"
    TestOnlyConfigureTestingBackend sbxPath
    m_TestingMode = True

    ' --- Act 1: getdb() in testing mode returns sandbox ---
    logs(3) = "4. Act: getdb() #1 with m_TestingMode=True — opens sbxPath"
    Set db1 = getdb(dbErr)
    If db1 Is Nothing Then
        Test_BackendCache_TestingModeToggleFromTrueToFalseReturnsCorrectBackend = JsonFail("getdb() #1 failed: " & dbErr, logs)
        GoTo Teardown
    End If
    If InStr(1, db1.Name, sbxPath, vbTextCompare) = 0 Then
        Test_BackendCache_TestingModeToggleFromTrueToFalseReturnsCorrectBackend = JsonFail("db1 not sbx: " & db1.Name, logs)
        GoTo Teardown
    End If

    ' --- Act 2: toggle m_TestingMode = False (no ResetTestSession) ---
    logs(4) = "5. Act: toggle m_TestingMode = False directly (NO ResetTestSession) — simulates caller bypassing the helper"
    m_TestingMode = False

    ' --- Act 3: second getdb() — should return prodPath, not stale sbxPath ---
    logs(5) = "6. Act: getdb() #2 — should detect URL mismatch and return " & prodPath & ", not cached sbxPath. RED on current code."
    Set db2 = getdb(dbErr)
    If db2 Is Nothing Then
        Test_BackendCache_TestingModeToggleFromTrueToFalseReturnsCorrectBackend = JsonFail("getdb() #2 failed: " & dbErr, logs)
        GoTo Teardown
    End If

    ' --- Assert ---
    logs(6) = "7. Assert: db2.Name must NOT contain " & sbxPath
    If InStr(1, db2.Name, sbxPath, vbTextCompare) > 0 Then
        Test_BackendCache_TestingModeToggleFromTrueToFalseReturnsCorrectBackend = JsonFail("STALE CACHE BUG: db2 still points to sbx after toggling m_TestingMode off", logs)
        GoTo Teardown
    End If

    logs(7) = "8. Assert: db2.Name must contain " & prodPath
    If InStr(1, db2.Name, prodPath, vbTextCompare) = 0 Then
        Test_BackendCache_TestingModeToggleFromTrueToFalseReturnsCorrectBackend = JsonFail("db2.Name does not match prod: " & db2.Name, logs)
        GoTo Teardown
    End If

    Test_BackendCache_TestingModeToggleFromTrueToFalseReturnsCorrectBackend = JsonOk("testing_mode_toggle_reopens_correct_backend", logs)

Teardown:
    On Error Resume Next
    Set db1 = Nothing
    Set db2 = Nothing
    m_TestingMode = False
    TestOnlyClearTestingBackend
    TestOnlyResetBackendConfigOverride
    CloseCachedBackendConnection
    DisposeIsolatedTempDb prodPath, dbTemp1
    DisposeIsolatedTempDb sbxPath, dbTemp2
    Exit Function

HandleError:
    Test_BackendCache_TestingModeToggleFromTrueToFalseReturnsCorrectBackend = JsonFail(Err.Description, logs)
    Resume Teardown
End Function

' ---- Test 5: AssertSandboxBackend must detect a stale cached connection ----

' Scenario under test (explore §3 / §6.1 Test 5):
'   1. m_TestingMode = True; m_TestingBackendURL = urlA; getdb() opens urlA
'   2. Re-point m_TestingBackendURL to urlB (without invalidating m_dbCached)
'   3. AssertSandboxBackend should detect that the cached connection
'      points to urlA while the active testing URL is urlB, and return False
'
' On the current code, AssertSandboxBackend only checks m_TestingMode,
' URL existence, and a "PROD" fingerprint — it does NOT inspect m_dbCached.
' So when the cached connection is stale, AssertSandboxBackend still
' returns True. This test is RED on the current implementation.
'
' Note: the GREEN fix in §6.2 (C1-C5) does NOT directly address this
' scenario. Test 5 documents the gap and will go GREEN when a future
' change adds a cache-staleness check to AssertSandboxBackend. Until
' that change lands, the test is RED.
'
' Schema/code evidence: TestHelper.bas lines 102-127 (AssertSandboxBackend)
' and Variables Globales.bas line 216 (m_dbCached private).
Public Function Test_BackendCache_AssertSandboxBackendDetectsStaleCache() As String
    Dim logs(0 To 7) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_BackendCache_AssertSandboxBackendDetectsStaleCache = JsonFail("test did not complete", logs)

    On Error GoTo HandleError
    ResetTestSession

    Dim errMsg As String
    Dim dbErr As String
    Dim urlA As String
    Dim urlB As String
    Dim dbTempA As DAO.Database
    Dim dbTempB As DAO.Database
    Dim dbCached As DAO.Database
    Dim assertErr As String
    Dim assertResult As Boolean

    ' --- Arrange ---
    logs(0) = "1. Arrange: create two temp .accdb backends for the sandbox URL change"
    urlA = MakeIsolatedTempDbPath("assert_a")
    Set dbTempA = CreateIsolatedTempDb(urlA, errMsg)
    If dbTempA Is Nothing Then
        Test_BackendCache_AssertSandboxBackendDetectsStaleCache = JsonFail("create urlA: " & errMsg, logs)
        GoTo Teardown
    End If
    Set dbTempA = Nothing

    urlB = MakeIsolatedTempDbPath("assert_b")
    Set dbTempB = CreateIsolatedTempDb(urlB, errMsg)
    If dbTempB Is Nothing Then
        Test_BackendCache_AssertSandboxBackendDetectsStaleCache = JsonFail("create urlB: " & errMsg, logs)
        GoTo Teardown
    End If
    Set dbTempB = Nothing

    logs(1) = "2. Arrange: configure testing backend to urlA, enable m_TestingMode"
    TestOnlyConfigureTestingBackend urlA
    m_TestingMode = True

    logs(2) = "3. Act: getdb() opens urlA — caches connection in m_dbCached"
    Set dbCached = getdb(dbErr)
    If dbCached Is Nothing Then
        Test_BackendCache_AssertSandboxBackendDetectsStaleCache = JsonFail("getdb() failed: " & dbErr, logs)
        GoTo Teardown
    End If
    If InStr(1, dbCached.Name, urlA, vbTextCompare) = 0 Then
        Test_BackendCache_AssertSandboxBackendDetectsStaleCache = JsonFail("getdb() not urlA: " & dbCached.Name, logs)
        GoTo Teardown
    End If

    ' --- Act: change sandbox URL without invalidating cache ---
    logs(3) = "4. Act: TestOnlyConfigureTestingBackend urlB → m_TestingBackendURL = " & urlB & " (cache still points to urlA)"
    TestOnlyConfigureTestingBackend urlB

    logs(4) = "5. Act: AssertSandboxBackend — should detect that m_dbCached points to urlA while testing URL is urlB (stale)"
    assertResult = AssertSandboxBackend(assertErr)

    ' --- Assert ---
    logs(5) = "6. Assert: AssertSandboxBackend must return False (m_dbCached is stale)"
    If assertResult Then
        Test_BackendCache_AssertSandboxBackendDetectsStaleCache = JsonFail("AssertSandboxBackend returned True but cache is stale (m_dbCached points to " & urlA & ", not " & urlB & "). Err=" & assertErr, logs)
        GoTo Teardown
    End If
    logs(6) = "7. Note: AssertSandboxBackend error message: " & assertErr

    Test_BackendCache_AssertSandboxBackendDetectsStaleCache = JsonOk("assert_sandbox_detects_stale_cache", logs)

Teardown:
    On Error Resume Next
    Set dbCached = Nothing
    m_TestingMode = False
    TestOnlyClearTestingBackend
    CloseCachedBackendConnection
    DisposeIsolatedTempDb urlA, dbTempA
    DisposeIsolatedTempDb urlB, dbTempB
    Exit Function

HandleError:
    Test_BackendCache_AssertSandboxBackendDetectsStaleCache = JsonFail(Err.Description, logs)
    Resume Teardown
End Function
