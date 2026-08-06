Attribute VB_Name = "Test_Getdb_Resolution"
Option Compare Database
Option Explicit

' ============================================================
' Test_Getdb_Resolution — Tests for align-getdb-no-conformidades-pattern
'
' Skill: access-vba-tdd v2.5 (HandleError pattern §1.1.2 + §1.7 isolation)
'
' Cobertura (5 tests atomicos, REQ-RES-001..005):
'   Test 1 - TestingMode_UsesSandbox (m_TestingMode=True con m_ActiveBackendURL seteado a path distinto
'           verifica que getdb usa m_BackendSandboxURL y no lee m_ActiveBackendURL) - GREEN pre-WU2
'   Test 2 - TestingMode_FailsWhenSandboxEmpty (m_TestingMode=True con m_BackendSandboxURL=""
'           verifica error especifico) - GREEN pre-WU2
'   Test 3 - NonTesting_UsesTempVars (m_TestingMode=False con TempVar seteado
'           verifica que getdb usa TempVar, no m_ActiveBackendURL) - RED pre-WU2
'   Test 4 - NonTesting_CallsLeeConfiguracionLocalWhenTempVarsEmpty (TempVar vacio
'           verifica que getdb llama LeeConfiguracionLocal y luego lee el TempVar) - RED pre-WU2
'   Test 5 - NonTesting_SkipConfigLoad_SkipsLeeConfiguracionLocal (p_SkipConfigLoad=True
'           verifica que getdb NO llama LeeConfiguracionLocal) - RED pre-WU2
' ============================================================

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

' --- State save/restore helpers (skill §1.7 isolation) ---
' Saves: m_TestingMode, m_BackendSandboxURL, m_ActiveBackendURL,
'        m_BackendActivo, Application.TempVars("BackendPathConfigurado")
' Returns a Variant array (indexed).
Private Function SaveBackendState() As Variant
    Dim state(0 To 4) As Variant
    state(0) = m_TestingMode
    state(1) = m_BackendSandboxURL
    state(2) = m_ActiveBackendURL
    state(3) = m_BackendActivo
    state(4) = Application.TempVars("BackendPathConfigurado")
    SaveBackendState = state
End Function

' Restores state from a save token.
' Per access-vba-tdd skill §1.7 (Inter-test isolation): this helper ALSO calls
' ResetGetDbCache (public wrapper in Variables Globales.bas) to ensure the next
' test starts with a clean getdb cache. Without this, a previous test's closed
' handle leaks into the next test's getdb() call, causing "Object is not valid"
' errors on the cache check in the non-testing branch.
Private Sub RestoreBackendState(ByRef p_State As Variant)
    On Error Resume Next
    ' 1. Reset getdb cache (closes m_CachedDB + resets m_DBOpen). Done FIRST so
    '    state restoration below cannot interfere with a stale handle.
    '    Module name has space -> brackets required: [Variables Globales].ResetGetDbCache
    [Variables Globales].ResetGetDbCache
    ' 2. Restore state values
    m_TestingMode = p_State(0)
    m_BackendSandboxURL = p_State(1)
    m_ActiveBackendURL = p_State(2)
    m_BackendActivo = p_State(3)
    If Len(Nz(p_State(4), "")) > 0 Then
        Application.TempVars("BackendPathConfigurado") = p_State(4)
    Else
        Application.TempVars.Remove "BackendPathConfigurado"
    End If
    On Error GoTo 0
End Sub

' Removes a TempVar key (idempotent).
Private Sub RemoveTempVar(ByVal p_Key As String)
    On Error Resume Next
    Application.TempVars.Remove p_Key
    On Error GoTo 0
End Sub

' ============================================================
' Test 1 - REQ-RES-001: getdb en testing mode usa m_BackendSandboxURL directo
' GIVEN m_TestingMode=True, m_BackendSandboxURL=valid local path
' WHEN getdb() is called
' THEN opens sandbox, NEVER reads m_ActiveBackendURL
' GREEN pre-WU2: testing branch unchanged in current code
' ============================================================
Public Function Test_Getdb_Resolution_TestingMode_UsesSandbox() As String
    Dim logs(0 To 6) As String
    Test_Getdb_Resolution_TestingMode_UsesSandbox = JsonFail("test did not complete", logs)
    On Error GoTo HandleError

    Dim cfgError As String
    Dim state As Variant
    state = SaveBackendState()
    logs(0) = "1. Arrange: EnsureTestConfigLoaded + state save"

    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_Getdb_Resolution_TestingMode_UsesSandbox = JsonFail(cfgError, logs)
        GoTo Teardown
    End If

    ' Verify preconditions
    If Not m_TestingMode Then
        Test_Getdb_Resolution_TestingMode_UsesSandbox = JsonFail("Precondition failed: m_TestingMode should be True after EnsureTestConfigLoaded", logs)
        GoTo Teardown
    End If
    If Len(Trim$(m_BackendSandboxURL)) = 0 Then
        Test_Getdb_Resolution_TestingMode_UsesSandbox = JsonFail("Precondition failed: m_BackendSandboxURL should be populated", logs)
        GoTo Teardown
    End If

    ' Pollute m_ActiveBackendURL with a non-sandbox value to verify getdb ignores it
    Dim originalActive As String
    originalActive = m_ActiveBackendURL
    m_ActiveBackendURL = "\\fake\prod\backend.accdb"
    logs(1) = "2. Arrange: m_ActiveBackendURL polluted with fake path"

    ' Act
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = getdb(dbErr)
    logs(2) = "3. Act: getdb() called"

    ' Assert: db is opened at the sandbox URL, not the polluted m_ActiveBackendURL
    If db Is Nothing Then
        Test_Getdb_Resolution_TestingMode_UsesSandbox = JsonFail("getdb returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Dim openedAt As String
    openedAt = db.Name
    If StrComp(openedAt, m_BackendSandboxURL, vbTextCompare) <> 0 Then
        Test_Getdb_Resolution_TestingMode_UsesSandbox = JsonFail("getdb opened at '" & openedAt & "' but expected sandbox '" & m_BackendSandboxURL & "'", logs)
        db.Close
        Set db = Nothing
        GoTo Teardown
    End If
    logs(3) = "4. Assert: getdb opened at sandbox path"

    ' Assert: m_ActiveBackendURL was NOT read or modified (still polluted)
    If m_ActiveBackendURL <> "\\fake\prod\backend.accdb" Then
        Test_Getdb_Resolution_TestingMode_UsesSandbox = JsonFail("m_ActiveBackendURL was modified by getdb: '" & m_ActiveBackendURL & "'", logs)
        db.Close
        Set db = Nothing
        GoTo Teardown
    End If
    logs(4) = "5. Assert: m_ActiveBackendURL untouched"

    db.Close
    Set db = Nothing
    Test_Getdb_Resolution_TestingMode_UsesSandbox = JsonOk("testing_mode_uses_sandbox_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    m_ActiveBackendURL = originalActive
    RestoreBackendState state
    On Error GoTo 0
    Exit Function

HandleError:
    Test_Getdb_Resolution_TestingMode_UsesSandbox = JsonFail("unexpected: " & Err.Description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' Test 2 - REQ-RES-001: testing mode con m_BackendSandboxURL vacio falla con error especifico
' GIVEN m_TestingMode=True, m_BackendSandboxURL=""
' WHEN getdb() is called
' THEN returns Nothing, p_Error contains "m_BackendSandboxURL"
' AND does NOT fall back to m_ActiveBackendURL
' GREEN pre-WU2: testing branch already returns this error
' ============================================================
Public Function Test_Getdb_Resolution_TestingMode_FailsWhenSandboxEmpty() As String
    Dim logs(0 To 5) As String
    Test_Getdb_Resolution_TestingMode_FailsWhenSandboxEmpty = JsonFail("test did not complete", logs)
    On Error GoTo HandleError

    Dim cfgError As String
    Dim state As Variant
    state = SaveBackendState()
    logs(0) = "1. Arrange: state save + EnsureTestConfigLoaded"

    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_Getdb_Resolution_TestingMode_FailsWhenSandboxEmpty = JsonFail(cfgError, logs)
        GoTo Teardown
    End If

    ' Force m_BackendSandboxURL empty
    Dim originalSandbox As String
    originalSandbox = m_BackendSandboxURL
    m_BackendSandboxURL = ""
    logs(1) = "2. Arrange: m_BackendSandboxURL forced empty"

    ' Pollute m_ActiveBackendURL with a valid-looking path; verify getdb does NOT fall back
    Dim originalActive As String
    originalActive = m_ActiveBackendURL
    m_ActiveBackendURL = "C:\fake\prod.accdb"
    logs(2) = "3. Arrange: m_ActiveBackendURL polluted (should not be used)"

    ' Act
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = getdb(dbErr)
    logs(3) = "4. Act: getdb() called"

    ' Assert: db is Nothing
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
        Test_Getdb_Resolution_TestingMode_FailsWhenSandboxEmpty = JsonFail("Expected db Is Nothing when sandbox is empty; got a valid handle", logs)
        GoTo Teardown
    End If

    ' Assert: p_Error contains the specific message
    If InStr(dbErr, "m_BackendSandboxURL") = 0 Then
        Test_Getdb_Resolution_TestingMode_FailsWhenSandboxEmpty = JsonFail("Expected p_Error to mention 'm_BackendSandboxURL'; got: '" & dbErr & "'", logs)
        GoTo Teardown
    End If
    logs(4) = "5. Assert: p_Error contains 'm_BackendSandboxURL'"

    Test_Getdb_Resolution_TestingMode_FailsWhenSandboxEmpty = JsonOk("testing_mode_fails_when_sandbox_empty_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    m_BackendSandboxURL = originalSandbox
    m_ActiveBackendURL = originalActive
    RestoreBackendState state
    On Error GoTo 0
    Exit Function

HandleError:
    Test_Getdb_Resolution_TestingMode_FailsWhenSandboxEmpty = JsonFail("unexpected: " & Err.Description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' Test 3 - REQ-RES-002: getdb en non-testing mode usa Application.TempVars("BackendPathConfigurado")
' GIVEN m_TestingMode=False, TempVar populated, m_ActiveBackendURL polluted
' WHEN getdb(p_SkipConfigLoad:=True) is called
' THEN opens TempVar path, NOT m_ActiveBackendURL
' RED pre-WU2: current getdb reads m_ActiveBackendURL, ignores TempVar
' ============================================================
Public Function Test_Getdb_Resolution_NonTesting_UsesTempVars() As String
    Dim logs(0 To 6) As String
    Test_Getdb_Resolution_NonTesting_UsesTempVars = JsonFail("test did not complete", logs)
    On Error GoTo HandleError

    Dim cfgError As String
    Dim state As Variant
    state = SaveBackendState()
    logs(0) = "1. Arrange: state save + EnsureTestConfigLoaded (to resolve sandbox path)"

    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_Getdb_Resolution_NonTesting_UsesTempVars = JsonFail(cfgError, logs)
        GoTo Teardown
    End If

    ' Switch to non-testing mode
    m_TestingMode = False
    logs(1) = "2. Arrange: m_TestingMode = False"

    ' Set TempVar to the resolved sandbox path
    Application.TempVars("BackendPathConfigurado") = m_BackendSandboxURL
    logs(2) = "3. Arrange: TempVar('BackendPathConfigurado') = sandbox path"

    ' Pollute m_ActiveBackendURL with a different value
    Dim originalActive As String
    originalActive = m_ActiveBackendURL
    m_ActiveBackendURL = "C:\fake\prod.accdb"
    logs(3) = "4. Arrange: m_ActiveBackendURL polluted (should be ignored)"

    ' Act with p_SkipConfigLoad=True so we isolate the TempVar read
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = getdb(dbErr, True)
    logs(4) = "5. Act: getdb(p_SkipConfigLoad:=True) called"

    ' Assert: db is not Nothing
    If db Is Nothing Then
        Test_Getdb_Resolution_NonTesting_UsesTempVars = JsonFail("getdb returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' Assert: db opened at the TempVar path
    If StrComp(db.Name, m_BackendSandboxURL, vbTextCompare) <> 0 Then
        Test_Getdb_Resolution_NonTesting_UsesTempVars = JsonFail("getdb opened at '" & db.Name & "' but expected TempVar path '" & m_BackendSandboxURL & "'", logs)
        db.Close
        Set db = Nothing
        GoTo Teardown
    End If
    logs(5) = "6. Assert: getdb opened at TempVar path"

    ' Assert: m_ActiveBackendURL was not read (still polluted)
    If m_ActiveBackendURL <> "C:\fake\prod.accdb" Then
        Test_Getdb_Resolution_NonTesting_UsesTempVars = JsonFail("m_ActiveBackendURL was modified by getdb", logs)
        db.Close
        Set db = Nothing
        GoTo Teardown
    End If
    logs(6) = "7. Assert: m_ActiveBackendURL untouched"

    db.Close
    Set db = Nothing
    Test_Getdb_Resolution_NonTesting_UsesTempVars = JsonOk("non_testing_uses_tempvars_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    m_ActiveBackendURL = originalActive
    RestoreBackendState state
    On Error GoTo 0
    Exit Function

HandleError:
    Test_Getdb_Resolution_NonTesting_UsesTempVars = JsonFail("unexpected: " & Err.Description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' Test 4 - REQ-RES-002: getdb llama LeeConfiguracionLocal cuando TempVar esta vacio
' GIVEN m_TestingMode=False, TempVar NOT set
' WHEN getdb() is called (p_SkipConfigLoad defaults to False)
' THEN LeeConfiguracionLocal is called, TempVar is populated, db opens
' RED pre-WU2: current getdb reads m_ActiveBackendURL, never calls LeeConfiguracionLocal on demand
' ============================================================
Public Function Test_Getdb_Resolution_NonTesting_CallsLeeConfiguracionLocalWhenTempVarsEmpty() As String
    Dim logs(0 To 5) As String
    Test_Getdb_Resolution_NonTesting_CallsLeeConfiguracionLocalWhenTempVarsEmpty = JsonFail("test did not complete", logs)
    On Error GoTo HandleError

    Dim cfgError As String
    Dim state As Variant
    state = SaveBackendState()
    logs(0) = "1. Arrange: state save + EnsureTestConfigLoaded"

    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_Getdb_Resolution_NonTesting_CallsLeeConfiguracionLocalWhenTempVarsEmpty = JsonFail(cfgError, logs)
        GoTo Teardown
    End If

    ' Switch to non-testing mode
    m_TestingMode = False

    ' Remove the TempVar to force LeeConfiguracionLocal to be called
    RemoveTempVar "BackendPathConfigurado"
    logs(1) = "2. Arrange: m_TestingMode=False, TempVar removed"

    ' Act: getdb() with default p_SkipConfigLoad=False
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = getdb(dbErr)
    logs(2) = "3. Act: getdb() called (default p_SkipConfigLoad=False)"

    ' Assert: LeeConfiguracionLocal populated the TempVar
    Dim tempAfter As String
    tempAfter = Nz(Application.TempVars("BackendPathConfigurado"), "")
    If Len(tempAfter) = 0 Then
        db.Close
        Set db = Nothing
        Test_Getdb_Resolution_NonTesting_CallsLeeConfiguracionLocalWhenTempVarsEmpty = JsonFail("After getdb(), TempVar('BackendPathConfigurado') is still empty (LeeConfiguracionLocal was NOT called or did not set it)", logs)
        GoTo Teardown
    End If
    logs(3) = "4. Assert: TempVar populated by LeeConfiguracionLocal"

    ' Assert: db is not Nothing (opened at the populated TempVar)
    If db Is Nothing Then
        Test_Getdb_Resolution_NonTesting_CallsLeeConfiguracionLocalWhenTempVarsEmpty = JsonFail("getdb returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' Assert: db opened at the TempVar path
    If StrComp(db.Name, tempAfter, vbTextCompare) <> 0 Then
        Test_Getdb_Resolution_NonTesting_CallsLeeConfiguracionLocalWhenTempVarsEmpty = JsonFail("getdb opened at '" & db.Name & "' but expected TempVar path '" & tempAfter & "'", logs)
        db.Close
        Set db = Nothing
        GoTo Teardown
    End If
    logs(4) = "5. Assert: getdb opened at TempVar path"

    db.Close
    Set db = Nothing
    Test_Getdb_Resolution_NonTesting_CallsLeeConfiguracionLocalWhenTempVarsEmpty = JsonOk("non_testing_calls_leeconfig_when_tempvar_empty_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    RestoreBackendState state
    On Error GoTo 0
    Exit Function

HandleError:
    Test_Getdb_Resolution_NonTesting_CallsLeeConfiguracionLocalWhenTempVarsEmpty = JsonFail("unexpected: " & Err.Description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' Test 5 - REQ-RES-003: p_SkipConfigLoad=True skips LeeConfiguracionLocal
' GIVEN m_TestingMode=False, TempVar NOT set
' WHEN getdb(p_SkipConfigLoad:=True) is called
' THEN returns Nothing, p_Error indicates "BackendPathConfigurado" not resolved
' AND LeeConfiguracionLocal is NOT called (TempVar still empty after)
' RED pre-WU2: current getdb has no p_SkipConfigLoad parameter, can't exercise this path
' ============================================================
Public Function Test_Getdb_Resolution_NonTesting_SkipConfigLoad_SkipsLeeConfiguracionLocal() As String
    Dim logs(0 To 5) As String
    Test_Getdb_Resolution_NonTesting_SkipConfigLoad_SkipsLeeConfiguracionLocal = JsonFail("test did not complete", logs)
    On Error GoTo HandleError

    Dim cfgError As String
    Dim state As Variant
    state = SaveBackendState()
    logs(0) = "1. Arrange: state save + EnsureTestConfigLoaded"

    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_Getdb_Resolution_NonTesting_SkipConfigLoad_SkipsLeeConfiguracionLocal = JsonFail(cfgError, logs)
        GoTo Teardown
    End If

    ' Switch to non-testing mode
    m_TestingMode = False

    ' Remove the TempVar
    RemoveTempVar "BackendPathConfigurado"
    logs(1) = "2. Arrange: m_TestingMode=False, TempVar removed"

    ' Act: getdb with p_SkipConfigLoad=True
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = getdb(dbErr, True)
    logs(2) = "3. Act: getdb(p_SkipConfigLoad:=True) called"

    ' Assert: db is Nothing
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
        Test_Getdb_Resolution_NonTesting_SkipConfigLoad_SkipsLeeConfiguracionLocal = JsonFail("Expected db Is Nothing when p_SkipConfigLoad=True and TempVar empty; got a valid handle", logs)
        GoTo Teardown
    End If

    ' Assert: p_Error mentions "BackendPathConfigurado"
    If InStr(dbErr, "BackendPathConfigurado") = 0 Then
        Test_Getdb_Resolution_NonTesting_SkipConfigLoad_SkipsLeeConfiguracionLocal = JsonFail("Expected p_Error to mention 'BackendPathConfigurado'; got: '" & dbErr & "'", logs)
        GoTo Teardown
    End If
    logs(3) = "4. Assert: p_Error mentions 'BackendPathConfigurado'"

    ' Assert: TempVar is still empty (LeeConfiguracionLocal was NOT called)
    If Len(Nz(Application.TempVars("BackendPathConfigurado"), "")) > 0 Then
        Test_Getdb_Resolution_NonTesting_SkipConfigLoad_SkipsLeeConfiguracionLocal = JsonFail("TempVar was populated despite p_SkipConfigLoad=True - LeeConfiguracionLocal was called incorrectly", logs)
        GoTo Teardown
    End If
    logs(4) = "5. Assert: TempVar still empty (LeeConfiguracionLocal NOT called)"

    Test_Getdb_Resolution_NonTesting_SkipConfigLoad_SkipsLeeConfiguracionLocal = JsonOk("non_testing_skipconfigload_skips_leeconfig_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    RestoreBackendState state
    On Error GoTo 0
    Exit Function

HandleError:
    Test_Getdb_Resolution_NonTesting_SkipConfigLoad_SkipsLeeConfiguracionLocal = JsonFail("unexpected: " & Err.Description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' RunAll - Aggregator (DOES NOT go in tests.vba.json per skill §4.7 manifest discipline)
' Lives here for ad-hoc smoke runs via dysflow_run_vba.
' ============================================================
Public Function Test_Getdb_Resolution_RunAll() As String
    Dim results(0 To 4) As String
    Dim names(0 To 4) As String
    Dim i As Long
    Dim allOk As Boolean
    Dim firstFailure As String
    Dim outLogs(0 To 6) As String

    names(0) = "Test_Getdb_Resolution_TestingMode_UsesSandbox"
    names(1) = "Test_Getdb_Resolution_TestingMode_FailsWhenSandboxEmpty"
    names(2) = "Test_Getdb_Resolution_NonTesting_UsesTempVars"
    names(3) = "Test_Getdb_Resolution_NonTesting_CallsLeeConfiguracionLocalWhenTempVarsEmpty"
    names(4) = "Test_Getdb_Resolution_NonTesting_SkipConfigLoad_SkipsLeeConfiguracionLocal"

    outLogs(0) = "RunAll: 5 atomic tests in sequence"
    results(0) = Test_Getdb_Resolution_TestingMode_UsesSandbox()
    results(1) = Test_Getdb_Resolution_TestingMode_FailsWhenSandboxEmpty()
    results(2) = Test_Getdb_Resolution_NonTesting_UsesTempVars()
    results(3) = Test_Getdb_Resolution_NonTesting_CallsLeeConfiguracionLocalWhenTempVarsEmpty()
    results(4) = Test_Getdb_Resolution_NonTesting_SkipConfigLoad_SkipsLeeConfiguracionLocal()

    allOk = True
    firstFailure = ""
    For i = 0 To 4
        If InStr(results(i), """ok"":false") > 0 Then
            allOk = False
            If firstFailure = "" Then firstFailure = names(i)
        End If
        outLogs(i + 1) = names(i) & ": " & IIf(InStr(results(i), """ok"":false") > 0, "FAIL", "OK")
    Next i

    If allOk Then
        Test_Getdb_Resolution_RunAll = JsonOk("all_pass", outLogs)
    Else
        Test_Getdb_Resolution_RunAll = JsonFail("first_failure: " & firstFailure, outLogs)
    End If
End Function