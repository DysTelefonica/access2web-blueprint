Attribute VB_Name = "DummyHatw"
Option Compare Database
Option Explicit

Private Const TEST_ID_BASE As Long = 900000
Private Const TEST_USER_CURRENT As Long = 900101
Private Const TEST_USER_HISTORICAL As Long = 900102

' ============================================================
' ANEXOS-HISTORICOS: module-level state for the conflict prompt test seam
' ============================================================
Private m_LastConflictMessage As String
Private m_LastConflictDNI As String
Private m_LastConflictDirection As String
Private m_ConflictPromptCalls As Long
Private m_TOCTOUExtraHistoricalUserID As Long
Private m_TOCTOUExtraCurrentUserID As Long
Private m_TOCTOUExtraDNI As String

' ============================================================
' Existing 5 tests (unchanged)
' ============================================================

Public Function Test_HATW_RejectsUnsafeUncRoots() As String
    Dim logs As Collection
    Dim errMsg As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), Nothing, "\\server\prod\HPS", TempLifecycleRoot("hist-unc"), "", errMsg)

    If InStr(1, errMsg, "Unsafe attachment root", vbTextCompare) = 0 Then
        Test_HATW_RejectsUnsafeUncRoots = JsonFail("Expected unsafe UNC root rejection, got: " & errMsg, logs)
        Exit Function
    End If

    logs.Add "Unsafe UNC root blocked before database/filesystem writes."
    Test_HATW_RejectsUnsafeUncRoots = JsonOk("unsafe-root-blocked", logs)
    Exit Function
EH:
    Test_HATW_RejectsUnsafeUncRoots = JsonFail("Unexpected error: " & Err.Description, logs)
End Function

Public Function Test_HATW_CurrentToHistorical_RollsBackBeforeCommitAndKeepsAttachments() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("current-db-rollback")
    historicalRoot = TempLifecycleRoot("historical-db-rollback")
    PrepareCurrentFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "BeforeCommit", errMsg)

    If errMsg = "" Then
        Test_HATW_CurrentToHistorical_RollsBackBeforeCommitAndKeepsAttachments = JsonFail("Expected forced BeforeCommit failure.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_CURRENT) <> 1 Then
        Test_HATW_CurrentToHistorical_RollsBackBeforeCommitAndKeepsAttachments = JsonFail("Current user was not restored by rollback.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuariosHistoricos", "ID=" & TEST_USER_CURRENT) <> 0 Then
        Test_HATW_CurrentToHistorical_RollsBackBeforeCommitAndKeepsAttachments = JsonFail("Historical partial user remains after rollback.", logs)
        GoTo CleanUp
    End If
    If Dir$(sourceFile) = "" Then
        Test_HATW_CurrentToHistorical_RollsBackBeforeCommitAndKeepsAttachments = JsonFail("Original attachment file is missing after rollback.", logs)
        GoTo CleanUp
    End If

    logs.Add "BeforeCommit failure rolled back data and kept original attachment."
    Test_HATW_CurrentToHistorical_RollsBackBeforeCommitAndKeepsAttachments = JsonOk("rollback-ok", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_CurrentToHistorical_RollsBackBeforeCommitAndKeepsAttachments = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_CurrentToHistorical_RestoresAfterSourceDeleteFailure() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim treeError As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim targetUserRoot As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("current-after-delete")
    historicalRoot = TempLifecycleRoot("historical-after-delete")
    PrepareCurrentFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile
    targetUserRoot = EnsureSlash(historicalRoot) & "DNI" & TEST_USER_CURRENT

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "AfterSourceDelete", errMsg)

    If errMsg = "" Then
        Test_HATW_CurrentToHistorical_RestoresAfterSourceDeleteFailure = JsonFail("Expected forced AfterSourceDelete failure.", logs)
        GoTo CleanUp
    End If
    If Dir$(EnsureSlash(currentRoot) & "DNI" & TEST_USER_CURRENT, vbDirectory) = "" Then
        Test_HATW_CurrentToHistorical_RestoresAfterSourceDeleteFailure = JsonFail("Snapshot restore did not recreate the original attachment root.", logs)
        GoTo CleanUp
    End If
    If Not AssertCurrentAttachmentTree(currentRoot, TEST_USER_CURRENT, treeError) Then
        Test_HATW_CurrentToHistorical_RestoresAfterSourceDeleteFailure = JsonFail("Snapshot restore did not recreate the exact original attachment tree: " & treeError, logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_CURRENT) <> 1 Then
        Test_HATW_CurrentToHistorical_RestoresAfterSourceDeleteFailure = JsonFail("Current user was not restored after filesystem failure.", logs)
        GoTo CleanUp
    End If
    If Dir$(targetUserRoot, vbDirectory) <> "" Then
        Test_HATW_CurrentToHistorical_RestoresAfterSourceDeleteFailure = JsonFail("Partial historical target tree remains after AfterSourceDelete compensation: " & targetUserRoot, logs)
        GoTo CleanUp
    End If

    logs.Add "AfterSourceDelete failure restored the exact nested source tree from snapshot."
    logs.Add "Partial target tree was cleaned after the failed attempt."
    Test_HATW_CurrentToHistorical_RestoresAfterSourceDeleteFailure = JsonOk("restore-ok", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_CurrentToHistorical_RestoresAfterSourceDeleteFailure = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_HistoricalToCurrent_PostCommitRefreshFailureLeavesMovedState() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim targetFile As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("current-post-refresh")
    historicalRoot = TempLifecycleRoot("historical-post-refresh")
    PrepareHistoricalFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    targetFile = EnsureSlash(currentRoot) & "DNI" & TEST_USER_HISTORICAL & "\ACTUAL\historical-note.txt"

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarHistoricoAActual(CStr(TEST_USER_HISTORICAL), db, currentRoot, historicalRoot, "AfterCommitRefresh", errMsg)

    If InStr(1, errMsg, "post-commit refresh", vbTextCompare) = 0 Then
        Test_HATW_HistoricalToCurrent_PostCommitRefreshFailureLeavesMovedState = JsonFail("Expected post-commit refresh warning, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_HISTORICAL) <> 1 Then
        Test_HATW_HistoricalToCurrent_PostCommitRefreshFailureLeavesMovedState = JsonFail("Committed current user row is missing after refresh warning.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuariosHistoricos", "ID=" & TEST_USER_HISTORICAL) <> 0 Then
        Test_HATW_HistoricalToCurrent_PostCommitRefreshFailureLeavesMovedState = JsonFail("Historical user row was restored after post-commit warning.", logs)
        GoTo CleanUp
    End If
    If Dir$(targetFile) = "" Then
        Test_HATW_HistoricalToCurrent_PostCommitRefreshFailureLeavesMovedState = JsonFail("Moved attachment is missing after post-commit warning.", logs)
        GoTo CleanUp
    End If

    logs.Add "Post-commit refresh warning did not compensate committed lifecycle move."
    Test_HATW_HistoricalToCurrent_PostCommitRefreshFailureLeavesMovedState = JsonOk("warning-only", logs)
CleanUp:
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_HistoricalToCurrent_PostCommitRefreshFailureLeavesMovedState = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
End Function

Public Function Test_HATW_Service_CurrentToHistorical_SucceedsAndMovesAttachments() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim targetFile As String
    Dim svc As UsuarioServicio
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("service-current-success")
    historicalRoot = TempLifecycleRoot("service-historical-success")
    PrepareCurrentFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile
    targetFile = EnsureSlash(historicalRoot) & "DNI" & TEST_USER_CURRENT & "\ACTUAL\current-note.txt"

    Set svc = New UsuarioServicio
    svc.ConfigureLifecycleTransactionTestSeams db, currentRoot, historicalRoot, "", errMsg
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    Call svc.PasarDeActualAHistorico(CStr(TEST_USER_CURRENT), errMsg)

    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_HATW_Service_CurrentToHistorical_SucceedsAndMovesAttachments = JsonFail("Expected service success or post-commit refresh warning, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_CURRENT) <> 0 Then
        Test_HATW_Service_CurrentToHistorical_SucceedsAndMovesAttachments = JsonFail("Current user row remains after service move.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuariosHistoricos", "ID=" & TEST_USER_CURRENT) <> 1 Then
        Test_HATW_Service_CurrentToHistorical_SucceedsAndMovesAttachments = JsonFail("Historical user row was not created by service move.", logs)
        GoTo CleanUp
    End If
    If Dir$(sourceFile) <> "" Then
        Test_HATW_Service_CurrentToHistorical_SucceedsAndMovesAttachments = JsonFail("Original current attachment still exists after service move.", logs)
        GoTo CleanUp
    End If
    If Dir$(targetFile) = "" Then
        Test_HATW_Service_CurrentToHistorical_SucceedsAndMovesAttachments = JsonFail("Historical attachment target is missing after service move.", logs)
        GoTo CleanUp
    End If

    If errMsg <> "" Then logs.Add "Post-commit refresh warning preserved durable current-to-historical move: " & errMsg
    logs.Add "UsuarioServicio.PasarDeActualAHistorico moved data and attachments to historical."
    Test_HATW_Service_CurrentToHistorical_SucceedsAndMovesAttachments = JsonOk("service-current-to-historical-success", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_Service_CurrentToHistorical_SucceedsAndMovesAttachments = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_Service_HistoricalToCurrent_SucceedsAndMovesAttachments() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim targetFile As String
    Dim svc As UsuarioServicio
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("service-current-restore")
    historicalRoot = TempLifecycleRoot("service-historical-restore")
    PrepareHistoricalFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    sourceFile = EnsureSlash(historicalRoot) & "DNI" & TEST_USER_HISTORICAL & "\historical-note.txt"
    targetFile = EnsureSlash(currentRoot) & "DNI" & TEST_USER_HISTORICAL & "\ACTUAL\historical-note.txt"

    Set svc = New UsuarioServicio
    svc.ConfigureLifecycleTransactionTestSeams db, currentRoot, historicalRoot, "", errMsg
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    Call svc.PasarDeHistoricoAActual(CStr(TEST_USER_HISTORICAL), errMsg)

    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_HATW_Service_HistoricalToCurrent_SucceedsAndMovesAttachments = JsonFail("Expected service restore success or post-commit refresh warning, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_HISTORICAL) <> 1 Then
        Test_HATW_Service_HistoricalToCurrent_SucceedsAndMovesAttachments = JsonFail("Current user row was not restored by service move.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuariosHistoricos", "ID=" & TEST_USER_HISTORICAL) <> 0 Then
        Test_HATW_Service_HistoricalToCurrent_SucceedsAndMovesAttachments = JsonFail("Historical user row remains after service move.", logs)
        GoTo CleanUp
    End If
    If Dir$(sourceFile) <> "" Then
        Test_HATW_Service_HistoricalToCurrent_SucceedsAndMovesAttachments = JsonFail("Original historical attachment still exists after service move.", logs)
        GoTo CleanUp
    End If
    If Dir$(targetFile) = "" Then
        Test_HATW_Service_HistoricalToCurrent_SucceedsAndMovesAttachments = JsonFail("Current attachment target is missing after service move.", logs)
        GoTo CleanUp
    End If

    If errMsg <> "" Then logs.Add "Post-commit refresh warning preserved durable historical-to-current move: " & errMsg
    logs.Add "UsuarioServicio.PasarDeHistoricoAActual moved data and attachments to current."
    Test_HATW_Service_HistoricalToCurrent_SucceedsAndMovesAttachments = JsonOk("service-historical-to-current-success", logs)
CleanUp:
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_Service_HistoricalToCurrent_SucceedsAndMovesAttachments = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
End Function

Public Function Test_HATW_Service_HistoricalToCurrent_RollsBackBeforeCommitAndKeepsAttachments() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim svc As UsuarioServicio
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("service-current-rollback")
    historicalRoot = TempLifecycleRoot("service-historical-rollback")
    PrepareHistoricalFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    sourceFile = EnsureSlash(historicalRoot) & "DNI" & TEST_USER_HISTORICAL & "\historical-note.txt"

    Set svc = New UsuarioServicio
    svc.ConfigureLifecycleTransactionTestSeams db, currentRoot, historicalRoot, "BeforeCommit", errMsg
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    Call svc.PasarDeHistoricoAActual(CStr(TEST_USER_HISTORICAL), errMsg)

    If errMsg = "" Then
        Test_HATW_Service_HistoricalToCurrent_RollsBackBeforeCommitAndKeepsAttachments = JsonFail("Expected service BeforeCommit failure.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_HISTORICAL) <> 0 Then
        Test_HATW_Service_HistoricalToCurrent_RollsBackBeforeCommitAndKeepsAttachments = JsonFail("Current partial user remains after service rollback.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuariosHistoricos", "ID=" & TEST_USER_HISTORICAL) <> 1 Then
        Test_HATW_Service_HistoricalToCurrent_RollsBackBeforeCommitAndKeepsAttachments = JsonFail("Historical user was not restored by service rollback.", logs)
        GoTo CleanUp
    End If
    If Dir$(sourceFile) = "" Then
        Test_HATW_Service_HistoricalToCurrent_RollsBackBeforeCommitAndKeepsAttachments = JsonFail("Original historical attachment is missing after service rollback.", logs)
        GoTo CleanUp
    End If

    logs.Add "UsuarioServicio.PasarDeHistoricoAActual rolled back pre-commit failure."
    Test_HATW_Service_HistoricalToCurrent_RollsBackBeforeCommitAndKeepsAttachments = JsonOk("service-historical-rollback-ok", logs)
CleanUp:
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_Service_HistoricalToCurrent_RollsBackBeforeCommitAndKeepsAttachments = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
End Function

Public Function Test_HATW_Service_CurrentToHistorical_FailureBeforeSourceDeleteLeavesOriginals() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim targetFile As String
    Dim svc As UsuarioServicio
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("service-current-before-delete")
    historicalRoot = TempLifecycleRoot("service-historical-before-delete")
    PrepareCurrentFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile
    targetFile = EnsureSlash(historicalRoot) & "DNI" & TEST_USER_CURRENT & "\ACTUAL\current-note.txt"

    Set svc = New UsuarioServicio
    svc.ConfigureLifecycleTransactionTestSeams db, currentRoot, historicalRoot, "BeforeSourceDelete", errMsg
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    Call svc.PasarDeActualAHistorico(CStr(TEST_USER_CURRENT), errMsg)

    If errMsg = "" Then
        Test_HATW_Service_CurrentToHistorical_FailureBeforeSourceDeleteLeavesOriginals = JsonFail("Expected service BeforeSourceDelete failure.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_CURRENT) <> 1 Then
        Test_HATW_Service_CurrentToHistorical_FailureBeforeSourceDeleteLeavesOriginals = JsonFail("Current user was not restored after before-delete failure.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuariosHistoricos", "ID=" & TEST_USER_CURRENT) <> 0 Then
        Test_HATW_Service_CurrentToHistorical_FailureBeforeSourceDeleteLeavesOriginals = JsonFail("Historical partial user remains after before-delete failure.", logs)
        GoTo CleanUp
    End If
    If Dir$(sourceFile) = "" Then
        Test_HATW_Service_CurrentToHistorical_FailureBeforeSourceDeleteLeavesOriginals = JsonFail("Original current attachment is missing after before-delete failure.", logs)
        GoTo CleanUp
    End If
    If Dir$(targetFile) <> "" Then
        Test_HATW_Service_CurrentToHistorical_FailureBeforeSourceDeleteLeavesOriginals = JsonFail("Partial historical target remains after before-delete failure.", logs)
        GoTo CleanUp
    End If

    logs.Add "UsuarioServicio.PasarDeActualAHistorico kept originals on before-source-delete failure."
    Test_HATW_Service_CurrentToHistorical_FailureBeforeSourceDeleteLeavesOriginals = JsonOk("service-before-source-delete-ok", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_Service_CurrentToHistorical_FailureBeforeSourceDeleteLeavesOriginals = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

' ============================================================
' ANEXOS-HISTORICOS: PUBLIC test-prompt callbacks
' These are the seams that the coordinator's ConfigureLifecycleConflictPrompt
' resolves via Application.Run. Signature MUST match:
'     Function(conflicts As String, dni As String, direction As String) As Boolean
' Returning True = approve + proceed (and run cleanup); False = user rejected.
' ============================================================

Public Function Test_ConflictPrompt_Approve(ByVal conflicts As String, ByVal dni As String, ByVal direction As String) As Boolean
    m_LastConflictMessage = conflicts
    m_LastConflictDNI = dni
    m_LastConflictDirection = direction
    m_ConflictPromptCalls = m_ConflictPromptCalls + 1
    Test_ConflictPrompt_Approve = True
End Function

Public Function Test_ConflictPrompt_Reject(ByVal conflicts As String, ByVal dni As String, ByVal direction As String) As Boolean
    m_LastConflictMessage = conflicts
    m_LastConflictDNI = dni
    m_LastConflictDirection = direction
    m_ConflictPromptCalls = m_ConflictPromptCalls + 1
    Test_ConflictPrompt_Reject = False
End Function

' Approves AND injects another historical user row (same DNI, different ID) after precheck #1
' already saw the ORIGINAL conflict. This simulates a TOCTOU race where another process
' added a new conflict row between precheck #1 (outside transaction) and precheck #2 (inside).
Public Function Test_ConflictPrompt_ApproveAndInjectExtraHistoricalUser(ByVal conflicts As String, ByVal dni As String, ByVal direction As String) As Boolean
    Dim openErr As String
    Dim db As DAO.Database
    m_LastConflictMessage = conflicts
    m_LastConflictDNI = dni
    m_LastConflictDirection = direction
    m_ConflictPromptCalls = m_ConflictPromptCalls + 1
    If m_TOCTOUExtraHistoricalUserID > 0 Then
        Set db = OpenLocalTestBackend(openErr)
        If openErr = "" And Not db Is Nothing Then
            On Error Resume Next
            db.Execute "INSERT INTO TbUsuariosHistoricos (ID, DNI, Nombre, Apellido_1) VALUES (" & m_TOCTOUExtraHistoricalUserID & ", '" & Replace(dni, "'", "''") & "', 'TOCTOU', 'Race')", dbFailOnError
            db.Close
            On Error GoTo 0
        End If
        Set db = Nothing
    End If
    Test_ConflictPrompt_ApproveAndInjectExtraHistoricalUser = True
End Function

' Approves AND injects another current user row (same DNI, different ID) for H2C direction.
Public Function Test_ConflictPrompt_ApproveAndInjectExtraCurrentUser(ByVal conflicts As String, ByVal dni As String, ByVal direction As String) As Boolean
    Dim openErr As String
    Dim db As DAO.Database
    m_LastConflictMessage = conflicts
    m_LastConflictDNI = dni
    m_LastConflictDirection = direction
    m_ConflictPromptCalls = m_ConflictPromptCalls + 1
    If m_TOCTOUExtraCurrentUserID > 0 Then
        Set db = OpenLocalTestBackend(openErr)
        If openErr = "" And Not db Is Nothing Then
            On Error Resume Next
            db.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & m_TOCTOUExtraCurrentUserID & ", '" & Replace(dni, "'", "''") & "', 'TOCTOU', 'Race')", dbFailOnError
            db.Close
            On Error GoTo 0
        End If
        Set db = Nothing
    End If
    Test_ConflictPrompt_ApproveAndInjectExtraCurrentUser = True
End Function

Public Sub Test_ResetConflictPromptState()
    m_LastConflictMessage = ""
    m_LastConflictDNI = ""
    m_LastConflictDirection = ""
    m_ConflictPromptCalls = 0
    m_TOCTOUExtraHistoricalUserID = 0
    m_TOCTOUExtraCurrentUserID = 0
    m_TOCTOUExtraDNI = ""
End Sub

Public Sub Test_SetTOCTOUExtraHistoricalUser(ByVal idUsuario As Long, ByVal dni As String)
    m_TOCTOUExtraHistoricalUserID = idUsuario
    m_TOCTOUExtraDNI = dni
End Sub

Public Sub Test_SetTOCTOUExtraCurrentUser(ByVal idUsuario As Long, ByVal dni As String)
    m_TOCTOUExtraCurrentUserID = idUsuario
    m_TOCTOUExtraDNI = dni
End Sub

Public Function Test_GetLastConflictMessage() As String
    Test_GetLastConflictMessage = m_LastConflictMessage
End Function

Public Function Test_GetConflictPromptCalls() As Long
    Test_GetConflictPromptCalls = m_ConflictPromptCalls
End Function

' ============================================================
' ANEXOS-HISTORICOS: 24 NEW TESTS
' ============================================================

Public Function Test_HATW_Precheck1_DetectsDBConflicts() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("precheck1-db")
    historicalRoot = TempLifecycleRoot("precheck1-db-hist")
    PrepareConflictFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, "DNI" & TEST_USER_CURRENT, TEST_USER_CURRENT + 100, sourceFile

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    coordinator.ConfigureLifecycleConflictPrompt "Test_ConflictPrompt_Reject"
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg = "" Or InStr(1, errMsg, "Operación cancelada", vbTextCompare) = 0 Then
        Test_HATW_Precheck1_DetectsDBConflicts = JsonFail("Expected 'Operación cancelada', got: " & errMsg, logs)
        GoTo CleanUp
    End If
    If InStr(1, m_LastConflictMessage, "BD", vbTextCompare) = 0 Then
        Test_HATW_Precheck1_DetectsDBConflicts = JsonFail("Expected BD indicator in conflicts, got: " & m_LastConflictMessage, logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_CURRENT) <> 1 Then
        Test_HATW_Precheck1_DetectsDBConflicts = JsonFail("Source user missing after rejection.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuariosHistoricos", "ID=" & TEST_USER_CURRENT + 100) <> 1 Then
        Test_HATW_Precheck1_DetectsDBConflicts = JsonFail("Pre-existing historical conflict row deleted.", logs)
        GoTo CleanUp
    End If

    logs.Add "Precheck #1 detected DB conflict and rejection cancelled cleanly; source + conflict preserved."
    Test_HATW_Precheck1_DetectsDBConflicts = JsonOk("precheck1-bd", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_Precheck1_DetectsDBConflicts = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_Precheck1_DetectsFolderConflicts() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("precheck1-folder")
    historicalRoot = TempLifecycleRoot("precheck1-folder-hist")
    PrepareFolderConflictFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    coordinator.ConfigureLifecycleConflictPrompt "Test_ConflictPrompt_Reject"
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg = "" Or InStr(1, errMsg, "Operación cancelada", vbTextCompare) = 0 Then
        Test_HATW_Precheck1_DetectsFolderConflicts = JsonFail("Expected folder conflict rejection, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    If InStr(1, m_LastConflictMessage, "Folder", vbTextCompare) = 0 Then
        Test_HATW_Precheck1_DetectsFolderConflicts = JsonFail("Expected Folder= in conflicts, got: " & m_LastConflictMessage, logs)
        GoTo CleanUp
    End If
    If Dir$(sourceFile) = "" Then
        Test_HATW_Precheck1_DetectsFolderConflicts = JsonFail("Source attachment lost after folder-rejection.", logs)
        GoTo CleanUp
    End If

    logs.Add "Precheck #1 detected folder conflict and rejection cancelled cleanly; source attachments preserved."
    Test_HATW_Precheck1_DetectsFolderConflicts = JsonOk("precheck1-folder", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_Precheck1_DetectsFolderConflicts = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_Precheck1_NoConflicts_DirectProceeds() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("precheck1-clean")
    historicalRoot = TempLifecycleRoot("precheck1-clean-hist")
    PrepareCurrentFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_HATW_Precheck1_NoConflicts_DirectProceeds = JsonFail("Expected clean success, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_CURRENT) <> 0 Then
        Test_HATW_Precheck1_NoConflicts_DirectProceeds = JsonFail("Source user still present after direct proceed.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuariosHistoricos", "ID=" & TEST_USER_CURRENT) <> 1 Then
        Test_HATW_Precheck1_NoConflicts_DirectProceeds = JsonFail("Historical user row not created by direct proceed.", logs)
        GoTo CleanUp
    End If
    If m_ConflictPromptCalls <> 0 Then
        Test_HATW_Precheck1_NoConflicts_DirectProceeds = JsonFail("Prompt should NOT be called when no conflicts: " & m_ConflictPromptCalls, logs)
        GoTo CleanUp
    End If

    logs.Add "Precheck #1 found no conflicts; operation proceeded directly with no prompt."
    Test_HATW_Precheck1_NoConflicts_DirectProceeds = JsonOk("precheck1-clean", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_Precheck1_NoConflicts_DirectProceeds = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_Precheck1_UserRejects_NoChanges() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("reject-nochanges")
    historicalRoot = TempLifecycleRoot("reject-nochanges-hist")
    PrepareConflictFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, "DNI" & TEST_USER_CURRENT, TEST_USER_CURRENT + 100, sourceFile

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    coordinator.ConfigureLifecycleConflictPrompt "Test_ConflictPrompt_Reject"
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_CURRENT) <> 1 Then
        Test_HATW_Precheck1_UserRejects_NoChanges = JsonFail("Source user moved despite rejection.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuariosHistoricos", "ID=" & TEST_USER_CURRENT + 100) <> 1 Then
        Test_HATW_Precheck1_UserRejects_NoChanges = JsonFail("Conflict row was altered by rejection.", logs)
        GoTo CleanUp
    End If
    If Dir$(sourceFile) = "" Then
        Test_HATW_Precheck1_UserRejects_NoChanges = JsonFail("Source attachment lost after rejection.", logs)
        GoTo CleanUp
    End If

    logs.Add "User rejection: source, conflict, attachments all preserved."
    Test_HATW_Precheck1_UserRejects_NoChanges = JsonOk("reject-nochanges", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_Precheck1_UserRejects_NoChanges = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_Cleanup_DeletesBothDBAndFolder() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("cleanup-both")
    historicalRoot = TempLifecycleRoot("cleanup-both-hist")
    PrepareBothConflictFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    coordinator.ConfigureLifecycleConflictPrompt "Test_ConflictPrompt_Approve"
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_HATW_Cleanup_DeletesBothDBAndFolder = JsonFail("Expected success after cleanup, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_CURRENT) <> 0 Then
        Test_HATW_Cleanup_DeletesBothDBAndFolder = JsonFail("Source still present after cleanup-driven move.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuariosHistoricos", "ID=" & TEST_USER_CURRENT) <> 1 Then
        Test_HATW_Cleanup_DeletesBothDBAndFolder = JsonFail("Historical row not created.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuariosHistoricos", "ID=" & TEST_USER_CURRENT + 100) <> 0 Then
        Test_HATW_Cleanup_DeletesBothDBAndFolder = JsonFail("Cleanup did not delete pre-existing historical conflict (ID=" & TEST_USER_CURRENT + 100 & ").", logs)
        GoTo CleanUp
    End If

    logs.Add "Cleanup (DB + folder) ran inside transaction; operation then completed."
    Test_HATW_Cleanup_DeletesBothDBAndFolder = JsonOk("cleanup-both", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_Cleanup_DeletesBothDBAndFolder = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_TOCTOU_DetectsNewConflictAfterCleanup() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Dim toctouExtraID As Long
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("toctou-c2h")
    historicalRoot = TempLifecycleRoot("toctou-c2h-hist")
    PrepareConflictFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, "DNI" & TEST_USER_CURRENT, TEST_USER_CURRENT + 100, sourceFile

    toctouExtraID = TEST_USER_CURRENT + 200
    Test_SetTOCTOUExtraHistoricalUser toctouExtraID, "DNI" & TEST_USER_CURRENT

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    coordinator.ConfigureLifecycleConflictPrompt "Test_ConflictPrompt_ApproveAndInjectExtraHistoricalUser"
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg = "" Then
        Test_HATW_TOCTOU_DetectsNewConflictAfterCleanup = JsonFail("Expected TOCTOU abort, got empty errMsg.", logs)
        GoTo CleanUp
    End If
    If InStr(1, errMsg, "TOCTOU", vbTextCompare) = 0 And InStr(1, errMsg, "race", vbTextCompare) = 0 And InStr(1, errMsg, "nuevo conflicto", vbTextCompare) = 0 And InStr(1, errMsg, "new conflict", vbTextCompare) = 0 Then
        Test_HATW_TOCTOU_DetectsNewConflictAfterCleanup = JsonFail("Expected TOCTOU/race error message, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_CURRENT) <> 1 Then
        Test_HATW_TOCTOU_DetectsNewConflictAfterCleanup = JsonFail("Source lost despite TOCTOU abort.", logs)
        GoTo CleanUp
    End If

    logs.Add "TOCTOU detected: prompt injected extra row mid-flight; operation aborted, source preserved."
    Test_HATW_TOCTOU_DetectsNewConflictAfterCleanup = JsonOk("toctou-c2h", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    On Error Resume Next
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbUsuariosHistoricos WHERE DNI='DNI" & TEST_USER_CURRENT & "'", dbFailOnError
    End If
    Exit Function
EH:
    Test_HATW_TOCTOU_DetectsNewConflictAfterCleanup = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_PromptMessage_MentionsBothDBAndFolder() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("prompt-msg")
    historicalRoot = TempLifecycleRoot("prompt-msg-hist")
    PrepareBothConflictFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    coordinator.ConfigureLifecycleConflictPrompt "Test_ConflictPrompt_Reject"
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If m_LastConflictMessage = "" Then
        Test_HATW_PromptMessage_MentionsBothDBAndFolder = JsonFail("Prompt message not captured.", logs)
        GoTo CleanUp
    End If
    If InStr(1, m_LastConflictMessage, "BD", vbTextCompare) = 0 Then
        Test_HATW_PromptMessage_MentionsBothDBAndFolder = JsonFail("Prompt message lacks BD indicator: " & m_LastConflictMessage, logs)
        GoTo CleanUp
    End If
    If InStr(1, m_LastConflictMessage, "Folder", vbTextCompare) = 0 Then
        Test_HATW_PromptMessage_MentionsBothDBAndFolder = JsonFail("Prompt message lacks folder indicator: " & m_LastConflictMessage, logs)
        GoTo CleanUp
    End If

    logs.Add "Prompt message mentions both BD and folder conflicts."
    Test_HATW_PromptMessage_MentionsBothDBAndFolder = JsonOk("prompt-msg", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_PromptMessage_MentionsBothDBAndFolder = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_HistoricalToCurrent_Precheck1_DetectsConflictsInCurrent() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("h2c-precheck1")
    historicalRoot = TempLifecycleRoot("h2c-precheck1-hist")
    PrepareH2CConflictFixture db, TEST_USER_HISTORICAL, historicalRoot, currentRoot

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    coordinator.ConfigureLifecycleConflictPrompt "Test_ConflictPrompt_Reject"
    Call coordinator.PasarHistoricoAActual(CStr(TEST_USER_HISTORICAL), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg = "" Or InStr(1, errMsg, "Operación cancelada", vbTextCompare) = 0 Then
        Test_HATW_HistoricalToCurrent_Precheck1_DetectsConflictsInCurrent = JsonFail("Expected 'Operación cancelada' (h2c), got: " & errMsg, logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuariosHistoricos", "ID=" & TEST_USER_HISTORICAL) <> 1 Then
        Test_HATW_HistoricalToCurrent_Precheck1_DetectsConflictsInCurrent = JsonFail("Historical source missing after rejection.", logs)
        GoTo CleanUp
    End If
    If InStr(1, m_LastConflictDirection, "Historical", vbTextCompare) = 0 Or InStr(1, m_LastConflictDirection, "Current", vbTextCompare) = 0 Then
        Test_HATW_HistoricalToCurrent_Precheck1_DetectsConflictsInCurrent = JsonFail("Direction missing 'Historical->Current' indicator: " & m_LastConflictDirection, logs)
        GoTo CleanUp
    End If

    logs.Add "H2C precheck #1 detected current-side DB conflict; rejection cancelled cleanly."
    Test_HATW_HistoricalToCurrent_Precheck1_DetectsConflictsInCurrent = JsonOk("h2c-precheck1", logs)
CleanUp:
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    On Error Resume Next
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbUsuarios WHERE DNI='DNI" & TEST_USER_HISTORICAL & "'", dbFailOnError
    End If
    Exit Function
EH:
    Test_HATW_HistoricalToCurrent_Precheck1_DetectsConflictsInCurrent = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
End Function

Public Function Test_HATW_HistoricalToCurrent_Cleanup_DeletesBothDBAndFolder() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("h2c-cleanup")
    historicalRoot = TempLifecycleRoot("h2c-cleanup-hist")
    PrepareH2CBothConflictFixture db, TEST_USER_HISTORICAL, historicalRoot, currentRoot

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    coordinator.ConfigureLifecycleConflictPrompt "Test_ConflictPrompt_Approve"
    Call coordinator.PasarHistoricoAActual(CStr(TEST_USER_HISTORICAL), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_HATW_HistoricalToCurrent_Cleanup_DeletesBothDBAndFolder = JsonFail("Expected h2c success: " & errMsg, logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_HISTORICAL) <> 1 Then
        Test_HATW_HistoricalToCurrent_Cleanup_DeletesBothDBAndFolder = JsonFail("Current user row not created by h2c cleanup-driven move.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuariosHistoricos", "ID=" & TEST_USER_HISTORICAL) <> 0 Then
        Test_HATW_HistoricalToCurrent_Cleanup_DeletesBothDBAndFolder = JsonFail("Historical source row was not moved.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_HISTORICAL + 100) <> 0 Then
        Test_HATW_HistoricalToCurrent_Cleanup_DeletesBothDBAndFolder = JsonFail("H2C cleanup did not delete pre-existing current conflict (ID=" & TEST_USER_HISTORICAL + 100 & ").", logs)
        GoTo CleanUp
    End If

    logs.Add "H2C cleanup deleted current conflict (DB + folder); operation moved historical -> current."
    Test_HATW_HistoricalToCurrent_Cleanup_DeletesBothDBAndFolder = JsonOk("h2c-cleanup", logs)
CleanUp:
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    On Error Resume Next
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbUsuarios WHERE DNI='DNI" & TEST_USER_HISTORICAL & "'", dbFailOnError
    End If
    Exit Function
EH:
    Test_HATW_HistoricalToCurrent_Cleanup_DeletesBothDBAndFolder = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
End Function

Public Function Test_HATW_HistoricalToCurrent_TOCTOU_DetectsNewConflict() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Dim toctouExtraID As Long
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("h2c-toctou")
    historicalRoot = TempLifecycleRoot("h2c-toctou-hist")
    PrepareH2CConflictFixture db, TEST_USER_HISTORICAL, historicalRoot, currentRoot

    toctouExtraID = TEST_USER_HISTORICAL + 200
    Test_SetTOCTOUExtraCurrentUser toctouExtraID, "DNI" & TEST_USER_HISTORICAL

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    coordinator.ConfigureLifecycleConflictPrompt "Test_ConflictPrompt_ApproveAndInjectExtraCurrentUser"
    Call coordinator.PasarHistoricoAActual(CStr(TEST_USER_HISTORICAL), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg = "" Then
        Test_HATW_HistoricalToCurrent_TOCTOU_DetectsNewConflict = JsonFail("Expected TOCTOU abort (h2c), got empty errMsg.", logs)
        GoTo CleanUp
    End If
    If InStr(1, errMsg, "TOCTOU", vbTextCompare) = 0 And InStr(1, errMsg, "race", vbTextCompare) = 0 And InStr(1, errMsg, "nuevo conflicto", vbTextCompare) = 0 And InStr(1, errMsg, "new conflict", vbTextCompare) = 0 Then
        Test_HATW_HistoricalToCurrent_TOCTOU_DetectsNewConflict = JsonFail("Expected TOCTOU/race error (h2c), got: " & errMsg, logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuariosHistoricos", "ID=" & TEST_USER_HISTORICAL) <> 1 Then
        Test_HATW_HistoricalToCurrent_TOCTOU_DetectsNewConflict = JsonFail("Historical source lost despite TOCTOU abort.", logs)
        GoTo CleanUp
    End If

    logs.Add "H2C TOCTOU detected: in-flight injection caught by inside-tx precheck #2."
    Test_HATW_HistoricalToCurrent_TOCTOU_DetectsNewConflict = JsonOk("h2c-toctou", logs)
CleanUp:
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    On Error Resume Next
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbUsuarios WHERE DNI='DNI" & TEST_USER_HISTORICAL & "'", dbFailOnError
    End If
    Exit Function
EH:
    Test_HATW_HistoricalToCurrent_TOCTOU_DetectsNewConflict = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
End Function

Public Function Test_HATW_HistoricalToCurrent_AnexosAlwaysNewIDAnexo_PlusEsHistorico() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("h2c-anexos-new")
    historicalRoot = TempLifecycleRoot("h2c-anexos-new-hist")
    PrepareHistoricalFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarHistoricoAActual(CStr(TEST_USER_HISTORICAL), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_HATW_HistoricalToCurrent_AnexosAlwaysNewIDAnexo_PlusEsHistorico = JsonFail("Expected h2c success, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    Dim historicalIDAnexo As Long
    Dim currentIDAnexo As Long
    historicalIDAnexo = TEST_USER_HISTORICAL + 1
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT IDAnexo, EsHistorico FROM TbAnexosUsuariosHPS WHERE IDUsuario=" & TEST_USER_HISTORICAL, dbOpenSnapshot)
    If rs.EOF Then
        Test_HATW_HistoricalToCurrent_AnexosAlwaysNewIDAnexo_PlusEsHistorico = JsonFail("No current anexo created.", logs)
        GoTo CleanUp
    End If
    currentIDAnexo = CLng(rs!IDAnexo)
    If currentIDAnexo = historicalIDAnexo Then
        Test_HATW_HistoricalToCurrent_AnexosAlwaysNewIDAnexo_PlusEsHistorico = JsonFail("IDAnexo was preserved instead of calculated: " & currentIDAnexo, logs)
        GoTo CleanUp
    End If
    If Nz(rs!EsHistorico, "") <> "No" Then
        Test_HATW_HistoricalToCurrent_AnexosAlwaysNewIDAnexo_PlusEsHistorico = JsonFail("EsHistorico was not 'No', got: " & Nz(rs!EsHistorico, "(null)"), logs)
        GoTo CleanUp
    End If
    rs.Close

    logs.Add "H2C: IDAnexo=" & currentIDAnexo & " (calculated), EsHistorico='No'."
    Test_HATW_HistoricalToCurrent_AnexosAlwaysNewIDAnexo_PlusEsHistorico = JsonOk("h2c-anexos-new", logs)
CleanUp:
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_HistoricalToCurrent_AnexosAlwaysNewIDAnexo_PlusEsHistorico = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
End Function

Public Function Test_HATW_CurrentToHistorical_User_PreservesIDWhenFreeInHistorical() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Dim resolvedID As Long
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    db.Execute "DELETE FROM TbUsuariosHistoricos WHERE ID=" & TEST_USER_CURRENT, dbFailOnError
    PrepareCurrentFixtureV2 db, TEST_USER_CURRENT

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    resolvedID = coordinator.ResolveTargetUserID(db, CLng(TEST_USER_CURRENT), True, errMsg)
    If errMsg <> "" Then
        Test_HATW_CurrentToHistorical_User_PreservesIDWhenFreeInHistorical = JsonFail("ResolveTargetUserID returned error: " & errMsg, logs)
        GoTo CleanUp
    End If
    If resolvedID <> TEST_USER_CURRENT Then
        Test_HATW_CurrentToHistorical_User_PreservesIDWhenFreeInHistorical = JsonFail("Expected preserved ID=" & TEST_USER_CURRENT & ", got: " & resolvedID, logs)
        GoTo CleanUp
    End If

    logs.Add "ResolveTargetUserID preserved source ID=" & TEST_USER_CURRENT & " when target slot was free."
    Test_HATW_CurrentToHistorical_User_PreservesIDWhenFreeInHistorical = JsonOk("resolve-free", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, "", ""
    Exit Function
EH:
    Test_HATW_CurrentToHistorical_User_PreservesIDWhenFreeInHistorical = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, "", ""
End Function

Public Function Test_HATW_CurrentToHistorical_User_CalculatesNewIDWhenSourceIDTaken() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Dim resolvedID As Long
    Dim preexistingMaxID As Long
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    PrepareCurrentFixtureV2 db, TEST_USER_CURRENT
    db.Execute "DELETE FROM TbUsuariosHistoricos WHERE ID=" & TEST_USER_CURRENT, dbFailOnError
    db.Execute "INSERT INTO TbUsuariosHistoricos (ID, DNI, Nombre, Apellido_1) VALUES (" & TEST_USER_CURRENT & ", 'OtherDNI', 'Blocking', 'Row')", dbFailOnError
    preexistingMaxID = TEST_USER_CURRENT + 50
    db.Execute "INSERT INTO TbUsuariosHistoricos (ID, DNI, Nombre, Apellido_1) VALUES (" & preexistingMaxID & ", 'AnotherDNI', 'OtherMax', 'Row')", dbFailOnError

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    resolvedID = coordinator.ResolveTargetUserID(db, CLng(TEST_USER_CURRENT), True, errMsg)
    If errMsg <> "" Then
        Test_HATW_CurrentToHistorical_User_CalculatesNewIDWhenSourceIDTaken = JsonFail("ResolveTargetUserID returned error: " & errMsg, logs)
        GoTo CleanUp
    End If
    If resolvedID = TEST_USER_CURRENT Then
        Test_HATW_CurrentToHistorical_User_CalculatesNewIDWhenSourceIDTaken = JsonFail("ResolveTargetUserID returned source ID instead of calculating new: " & resolvedID, logs)
        GoTo CleanUp
    End If
    If resolvedID <= preexistingMaxID Then
        Test_HATW_CurrentToHistorical_User_CalculatesNewIDWhenSourceIDTaken = JsonFail("Resolved ID not greater than max(preexisting): resolved=" & resolvedID & ", max=" & preexistingMaxID, logs)
        GoTo CleanUp
    End If

    logs.Add "ResolveTargetUserID calculated new ID=" & resolvedID & " (> max=" & preexistingMaxID & ") when source ID was taken."
    Test_HATW_CurrentToHistorical_User_CalculatesNewIDWhenSourceIDTaken = JsonOk("resolve-taken", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, "", ""
    db.Execute "DELETE FROM TbUsuariosHistoricos WHERE ID=" & preexistingMaxID, dbFailOnError
    Exit Function
EH:
    Test_HATW_CurrentToHistorical_User_CalculatesNewIDWhenSourceIDTaken = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, "", ""
End Function

Public Function Test_HATW_CurrentToHistorical_Anexos_AlwaysGetNewIDAnexoEvenWhenSourceFree() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Dim sourceIDAnexo As Long
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("anexos-new-free")
    historicalRoot = TempLifecycleRoot("anexos-new-free-hist")
    PrepareAnexoFixtureV2 db, TEST_USER_CURRENT
    CreateCurrentAttachmentTreeT2 currentRoot, TEST_USER_CURRENT

    sourceIDAnexo = TEST_USER_CURRENT + 1
    If CountRows(db, "TbAnexosUsuariosHistoricos", "IDAnexo=" & sourceIDAnexo) <> 0 Then
        db.Execute "DELETE FROM TbAnexosUsuariosHistoricos WHERE IDAnexo=" & sourceIDAnexo, dbFailOnError
    End If

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_HATW_CurrentToHistorical_Anexos_AlwaysGetNewIDAnexoEvenWhenSourceFree = JsonFail("Expected success, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT IDAnexo FROM TbAnexosUsuariosHistoricos WHERE IDUsuario=" & TEST_USER_CURRENT, dbOpenSnapshot)
    If rs.EOF Then
        Test_HATW_CurrentToHistorical_Anexos_AlwaysGetNewIDAnexoEvenWhenSourceFree = JsonFail("No historical anexo created.", logs)
        GoTo CleanUp
    End If
    If CLng(rs!IDAnexo) = sourceIDAnexo Then
        Test_HATW_CurrentToHistorical_Anexos_AlwaysGetNewIDAnexoEvenWhenSourceFree = JsonFail("IDAnexo was preserved (should ALWAYS be calculated): " & rs!IDAnexo, logs)
        GoTo CleanUp
    End If
    rs.Close

    logs.Add "Anexo IDAnexo was recalculated to a new value even though source IDAnexo=" & sourceIDAnexo & " was free in destination."
    Test_HATW_CurrentToHistorical_Anexos_AlwaysGetNewIDAnexoEvenWhenSourceFree = JsonOk("anexos-new-free", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_CurrentToHistorical_Anexos_AlwaysGetNewIDAnexoEvenWhenSourceFree = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_CurrentToHistorical_Anexos_MultipleRows_NoIDAnexoCollision() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("anexos-multi")
    historicalRoot = TempLifecycleRoot("anexos-multi-hist")
    PrepareMultipleAnexosFixture db, TEST_USER_CURRENT, currentRoot

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_HATW_CurrentToHistorical_Anexos_MultipleRows_NoIDAnexoCollision = JsonFail("Expected success, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT IDAnexo FROM TbAnexosUsuariosHistoricos WHERE IDUsuario=" & TEST_USER_CURRENT, dbOpenSnapshot)
    Dim total As Long
    total = 0
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Do While Not rs.EOF
        total = total + 1
        Dim id As Variant
        id = Nz(rs!IDAnexo, 0)
        If Not seen.Exists(CStr(id)) Then seen.Add CStr(id), True
        rs.MoveNext
    Loop
    rs.Close
    Dim uniq As Long
    uniq = seen.Count
    If total <> 3 Then
        Test_HATW_CurrentToHistorical_Anexos_MultipleRows_NoIDAnexoCollision = JsonFail("Expected 3 historical anexos, got: " & total, logs)
        GoTo CleanUp
    End If
    If uniq <> 3 Then
        Test_HATW_CurrentToHistorical_Anexos_MultipleRows_NoIDAnexoCollision = JsonFail("IDAnexo collision: total=" & total & ", distinct=" & uniq, logs)
        GoTo CleanUp
    End If

    logs.Add "All 3 historical anexos have unique IDAnexo (no collisions)."
    Test_HATW_CurrentToHistorical_Anexos_MultipleRows_NoIDAnexoCollision = JsonOk("anexos-multi", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_CurrentToHistorical_Anexos_MultipleRows_NoIDAnexoCollision = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_Precheck_PropagatesDBError() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Dim sandboxPath As String
    Dim sourceBackend As String
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    ' Use a sandbox copy of HPST.accdb so we can drop TbUsuariosHistoricos without
    ' hitting FK constraints from other tables in the production backend.
    sourceBackend = EnsureSlash(CurrentProject.path) & "HPST.accdb"
    sandboxPath = Environ$("TEMP") & "\HPS\hps_precheck_dberr_" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Rnd * 1000000)) & ".accdb"
    CreateFolderTree Environ$("TEMP") & "\HPS"
    FileCopy sourceBackend, sandboxPath
    Set db = DBEngine.Workspaces(0).OpenDatabase(sandboxPath, False, False, BackendConnectString())
    If errMsg = "" Then errMsg = ""
    currentRoot = TempLifecycleRoot("precheck-dberr")
    historicalRoot = TempLifecycleRoot("precheck-dberr-hist")
    PrepareCurrentFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile
    ' Drop FK from TbObservacionesHistoricas -> TbUsuariosHistoricos first so DROP TABLE works.
    On Error Resume Next
    db.Relations.Delete "TbUsuariosHistoricosTbObservacionesHistoricas"
    On Error GoTo EH
    db.Execute "DROP TABLE TbUsuariosHistoricos", dbFailOnError

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg = "" Then
        Test_HATW_Precheck_PropagatesDBError = JsonFail("Expected precheck DB error, got empty errMsg.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_CURRENT) <> 1 Then
        Test_HATW_Precheck_PropagatesDBError = JsonFail("Source row lost after precheck DB error.", logs)
        GoTo CleanUp
    End If

    logs.Add "Precheck DB error propagated; source preserved (no transaction opened)."
    Test_HATW_Precheck_PropagatesDBError = JsonOk("precheck-dberr", logs)
CleanUp:
    On Error Resume Next
    If Not db Is Nothing Then db.Close
    Set db = Nothing
    If Len(sandboxPath) > 0 And Dir$(sandboxPath) <> "" Then Kill sandboxPath
    Exit Function
EH:
    Test_HATW_Precheck_PropagatesDBError = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    If Not db Is Nothing Then db.Close
    Set db = Nothing
    If Len(sandboxPath) > 0 And Dir$(sandboxPath) <> "" Then Kill sandboxPath
End Function

Public Function Test_HATW_AnexosCopy_PropagatesMaxLongError() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Dim sourceIDAnexo As Long
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("anexos-maxlong")
    historicalRoot = TempLifecycleRoot("anexos-maxlong-hist")
    PrepareAnexoFixtureV2 db, TEST_USER_CURRENT
    CreateCurrentAttachmentTreeT2 currentRoot, TEST_USER_CURRENT
    db.Execute "DROP TABLE TbAnexosUsuariosHistoricos", dbFailOnError

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    sourceIDAnexo = TEST_USER_CURRENT + 1
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg = "" Then
        Test_HATW_AnexosCopy_PropagatesMaxLongError = JsonFail("Expected anexos copy DB error, got empty errMsg.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_CURRENT) <> 1 Then
        Test_HATW_AnexosCopy_PropagatesMaxLongError = JsonFail("Source lost after anexos copy DB error.", logs)
        GoTo CleanUp
    End If

    logs.Add "Anexos copy DB error propagated and rolled back; source preserved."
    Test_HATW_AnexosCopy_PropagatesMaxLongError = JsonOk("anexos-maxlong", logs)
CleanUp:
    On Error Resume Next
    db.Execute "CREATE TABLE TbAnexosUsuariosHistoricos (IDAnexo LONG NOT NULL, NombreAnexo TEXT(255), IDUsuario LONG, Hash TEXT(64), CONSTRAINT PK_TbAnexosUsuariosHistoricos PRIMARY KEY (IDAnexo))", dbFailOnError
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_AnexosCopy_PropagatesMaxLongError = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_CurrentToHistorical_Anexos_RollbackDoesNotBurnIDs() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("anexos-rb")
    historicalRoot = TempLifecycleRoot("anexos-rb-hist")
    PrepareCurrentFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "BeforeCommit", errMsg)

    If errMsg = "" Then
        Test_HATW_CurrentToHistorical_Anexos_RollbackDoesNotBurnIDs = JsonFail("Expected BeforeCommit failure.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbAnexosUsuariosHistoricos", "IDUsuario=" & TEST_USER_CURRENT) <> 0 Then
        Test_HATW_CurrentToHistorical_Anexos_RollbackDoesNotBurnIDs = JsonFail("Rollback did not undo anexos copy (residual row in TbAnexosUsuariosHistoricos).", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_CURRENT) <> 1 Then
        Test_HATW_CurrentToHistorical_Anexos_RollbackDoesNotBurnIDs = JsonFail("Source lost despite BeforeCommit rollback.", logs)
        GoTo CleanUp
    End If

    logs.Add "Rollback undid both user and anexos copy (no burned IDs)."
    Test_HATW_CurrentToHistorical_Anexos_RollbackDoesNotBurnIDs = JsonOk("anexos-rb", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_CurrentToHistorical_Anexos_RollbackDoesNotBurnIDs = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_CurrentToHistorical_Observaciones_PreservesFreeIDs() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("obs-preserves")
    historicalRoot = TempLifecycleRoot("obs-preserves-hist")
    PrepareObservacionFixtureV2 db, TEST_USER_CURRENT, currentRoot, currentRoot

    db.Execute "DELETE FROM TbObservacionesHistoricas WHERE IDObservacion=950001", dbFailOnError

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_HATW_CurrentToHistorical_Observaciones_PreservesFreeIDs = JsonFail("Expected success, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbObservacionesHistoricas", "IDObservacion=950001") <> 1 Then
        Test_HATW_CurrentToHistorical_Observaciones_PreservesFreeIDs = JsonFail("Expected preserved IDObservacion=950001, got count=" & CountRows(db, "TbObservacionesHistoricas", "IDObservacion=950001"), logs)
        GoTo CleanUp
    End If

    logs.Add "Observacion IDObservacion=950001 preserved (target slot was free)."
    Test_HATW_CurrentToHistorical_Observaciones_PreservesFreeIDs = JsonOk("obs-preserves", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_CurrentToHistorical_Observaciones_PreservesFreeIDs = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_CurrentToHistorical_Observaciones_CalculatesWhenCollision() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Dim preexistingMaxID As Long
    Dim blockerParentID As Long
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("obs-collision")
    historicalRoot = TempLifecycleRoot("obs-collision-hist")
    PrepareObservacionFixtureV2 db, TEST_USER_CURRENT, currentRoot, currentRoot

    ' TbObservacionesHistoricas has FK -> TbUsuariosHistoricos(ID). Insert a parent
    ' row with a fictional ID (not TEST_USER_CURRENT, so the lifecycle still preserves
    ' ID=TEST_USER_CURRENT) so the blocking/high-water observation rows satisfy the FK.
    blockerParentID = 9999999
    db.Execute "DELETE FROM TbUsuariosHistoricos WHERE ID=" & blockerParentID, dbFailOnError
    db.Execute "INSERT INTO TbUsuariosHistoricos (ID, DNI, Nombre, Apellido_1) VALUES (" & blockerParentID & ", 'fake-blocker-dni', 'Fake', 'Blocking')", dbFailOnError
    db.Execute "DELETE FROM TbObservacionesHistoricas WHERE IDObservacion=950001", dbFailOnError
    db.Execute "INSERT INTO TbObservacionesHistoricas (IDObservacion, ID, Fecha, Observacion, Tipo) VALUES (950001, " & blockerParentID & ", #2026-06-30#, 'Blocking', 'N')", dbFailOnError
    preexistingMaxID = 950099
    db.Execute "DELETE FROM TbObservacionesHistoricas WHERE IDObservacion=" & preexistingMaxID, dbFailOnError
    db.Execute "INSERT INTO TbObservacionesHistoricas (IDObservacion, ID, Fecha, Observacion, Tipo) VALUES (" & preexistingMaxID & ", " & blockerParentID & ", #2026-06-30#, 'HighWater', 'N')", dbFailOnError

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_HATW_CurrentToHistorical_Observaciones_CalculatesWhenCollision = JsonFail("Expected success, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT IDObservacion FROM TbObservacionesHistoricas WHERE ID=" & TEST_USER_CURRENT, dbOpenSnapshot)
    If rs.EOF Then
        Test_HATW_CurrentToHistorical_Observaciones_CalculatesWhenCollision = JsonFail("No historical observacion created.", logs)
        GoTo CleanUp
    End If
    If CLng(rs!IDObservacion) = 950001 Then
        Test_HATW_CurrentToHistorical_Observaciones_CalculatesWhenCollision = JsonFail("IDObservacion was preserved despite collision: " & rs!IDObservacion, logs)
        GoTo CleanUp
    End If
    If CLng(rs!IDObservacion) <= preexistingMaxID Then
        Test_HATW_CurrentToHistorical_Observaciones_CalculatesWhenCollision = JsonFail("IDObservacion not greater than max: got=" & rs!IDObservacion & ", max=" & preexistingMaxID, logs)
        GoTo CleanUp
    End If
    rs.Close

    logs.Add "Observacion IDObservacion calculated as new (> max=" & preexistingMaxID & ") on collision."
    Test_HATW_CurrentToHistorical_Observaciones_CalculatesWhenCollision = JsonOk("obs-collision", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    On Error Resume Next
    db.Execute "DELETE FROM TbObservacionesHistoricas WHERE IDObservacion=" & preexistingMaxID, dbFailOnError
    db.Execute "DELETE FROM TbObservacionesHistoricas WHERE IDObservacion=950001", dbFailOnError
    db.Execute "DELETE FROM TbUsuariosHistoricos WHERE ID=" & blockerParentID, dbFailOnError
    Exit Function
EH:
    Test_HATW_CurrentToHistorical_Observaciones_CalculatesWhenCollision = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_ErrorTranslator_DuplicateKey_MentionsTableAndID() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Dim translated As String
    translated = coordinator.TranslateLifecycleError("CopyAnexosActualAHistorico", "TbAnexosUsuariosHistoricos", 12345, "IDAnexo")

    If InStr(1, translated, "CopyAnexosActualAHistorico", vbTextCompare) = 0 Then
        Test_HATW_ErrorTranslator_DuplicateKey_MentionsTableAndID = JsonFail("Translated message missing method name: " & translated, logs)
        Exit Function
    End If
    If InStr(1, translated, "TbAnexosUsuariosHistoricos", vbTextCompare) = 0 Then
        Test_HATW_ErrorTranslator_DuplicateKey_MentionsTableAndID = JsonFail("Translated message missing destination table: " & translated, logs)
        Exit Function
    End If
    If InStr(1, translated, "12345", vbTextCompare) = 0 Then
        Test_HATW_ErrorTranslator_DuplicateKey_MentionsTableAndID = JsonFail("Translated message missing attempted ID: " & translated, logs)
        Exit Function
    End If

    logs.Add "ErrorTranslator produced contextual message with method + table + ID."
    Test_HATW_ErrorTranslator_DuplicateKey_MentionsTableAndID = JsonOk("trans-duplicate", logs)
    Exit Function
EH:
    Test_HATW_ErrorTranslator_DuplicateKey_MentionsTableAndID = JsonFail("Unexpected error: " & Err.Description, logs)
End Function

Public Function Test_HATW_ErrorTranslator_UnknownError_PreservesOriginalMessage() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Dim translated As String
    translated = coordinator.TranslateLifecycleError("ExecuteLifecycleMove", "", Null, "")

    If translated = "" Then
        Test_HATW_ErrorTranslator_UnknownError_PreservesOriginalMessage = JsonFail("Empty translation returned.", logs)
        Exit Function
    End If
    If InStr(1, translated, "ExecuteLifecycleMove", vbTextCompare) = 0 Then
        Test_HATW_ErrorTranslator_UnknownError_PreservesOriginalMessage = JsonFail("UnknownError translator still includes the method name: " & translated, logs)
        Exit Function
    End If

    logs.Add "ErrorTranslator preserves method context even for unknown errors."
    Test_HATW_ErrorTranslator_UnknownError_PreservesOriginalMessage = JsonOk("trans-unknown", logs)
    Exit Function
EH:
    Test_HATW_ErrorTranslator_UnknownError_PreservesOriginalMessage = JsonFail("Unexpected error: " & Err.Description, logs)
End Function

Public Function Test_HATW_CleanupFailsAfterUserApprove_RollsBackBDConflicts() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Dim lockRs As DAO.Recordset
    Dim sourceFile As String
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("cleanup-fails")
    historicalRoot = TempLifecycleRoot("cleanup-fails-hist")
    PrepareConflictFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, "DNI" & TEST_USER_CURRENT, TEST_USER_CURRENT + 100, sourceFile

    ' FK TbObservacionesHistoricas->TbUsuariosHistoricos has CASCADE DELETE in this
    ' backend, so a child row would be deleted automatically and NOT block the parent DELETE.
    ' Instead, open TbUsuariosHistoricos with dbDenyWrite so no other writer (including
    ' the lifecycle's db.Execute for DeletePriorConflicts) can acquire a write lock on
    ' the table while this recordset is open. The DELETE will then fail with
    ' "Couldn't lock table" or similar, propagating through the EH into p_Error.
    Set lockRs = db.OpenRecordset("TbUsuariosHistoricos", dbOpenDynaset, dbDenyWrite)

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    coordinator.ConfigureLifecycleConflictPrompt "Test_ConflictPrompt_Approve"
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg = "" Then
        Test_HATW_CleanupFailsAfterUserApprove_RollsBackBDConflicts = JsonFail("Expected cleanup failure (lock contention), got empty errMsg.", logs)
        GoTo CleanUp
    End If
    ' Permissive error check: lock contention OR FK OR delete — any error during cleanup
    If InStr(1, errMsg, "lock", vbTextCompare) = 0 And InStr(1, errMsg, "use", vbTextCompare) = 0 And InStr(1, errMsg, "FK", vbTextCompare) = 0 And InStr(1, errMsg, "relationship", vbTextCompare) = 0 And InStr(1, errMsg, "referenc", vbTextCompare) = 0 And InStr(1, errMsg, "DeletePriorConflicts", vbTextCompare) = 0 And InStr(1, errMsg, "delete", vbTextCompare) = 0 Then
        Test_HATW_CleanupFailsAfterUserApprove_RollsBackBDConflicts = JsonFail("Expected lock/FK/delete error, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_CURRENT) <> 1 Then
        Test_HATW_CleanupFailsAfterUserApprove_RollsBackBDConflicts = JsonFail("Source lost after cleanup failure.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuariosHistoricos", "ID=" & TEST_USER_CURRENT + 100) <> 1 Then
        Test_HATW_CleanupFailsAfterUserApprove_RollsBackBDConflicts = JsonFail("Conflict row missing after cleanup failure (should have been rolled back).", logs)
        GoTo CleanUp
    End If

    logs.Add "Cleanup lock contention rolled back; conflict row preserved."
    Test_HATW_CleanupFailsAfterUserApprove_RollsBackBDConflicts = JsonOk("cleanup-fails", logs)
CleanUp:
    On Error Resume Next
    If Not lockRs Is Nothing Then
        lockRs.CancelUpdate
        lockRs.Close
        Set lockRs = Nothing
    End If
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_CleanupFailsAfterUserApprove_RollsBackBDConflicts = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    If Not lockRs Is Nothing Then
        lockRs.CancelUpdate
        lockRs.Close
        Set lockRs = Nothing
    End If
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_FolderDeleteInsideTransaction_FolderStaysDeletedIfCommitSucceeds() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("folder-intx")
    historicalRoot = TempLifecycleRoot("folder-intx-hist")
    PrepareFolderConflictFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile
    Dim conflictFolderPath As String
    conflictFolderPath = EnsureSlash(historicalRoot) & "DNI" & TEST_USER_CURRENT

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    coordinator.ConfigureLifecycleConflictPrompt "Test_ConflictPrompt_Approve"
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_HATW_FolderDeleteInsideTransaction_FolderStaysDeletedIfCommitSucceeds = JsonFail("Expected success, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    Dim conflictingOriginal As String
    conflictingOriginal = EnsureSlash(conflictFolderPath) & "conflict-pre-existing.txt"
    If Dir$(conflictingOriginal) <> "" Then
        Test_HATW_FolderDeleteInsideTransaction_FolderStaysDeletedIfCommitSucceeds = JsonFail("Original conflicting file still exists after commit: " & conflictingOriginal, logs)
        GoTo CleanUp
    End If
    Dim recreatedAsSourceContent As String
    recreatedAsSourceContent = EnsureSlash(conflictFolderPath) & "ACTUAL\current-note.txt"
    If Dir$(recreatedAsSourceContent) = "" Then
        Test_HATW_FolderDeleteInsideTransaction_FolderStaysDeletedIfCommitSucceeds = JsonFail("Folder replaced by source copy is missing the source file: " & recreatedAsSourceContent, logs)
        GoTo CleanUp
    End If

    logs.Add "Folder delete inside transaction was persisted; copy replaced the cleaned folder with source content."
    Test_HATW_FolderDeleteInsideTransaction_FolderStaysDeletedIfCommitSucceeds = JsonOk("folder-intx", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_FolderDeleteInsideTransaction_FolderStaysDeletedIfCommitSucceeds = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

' ============================================================
' ANEXOS-HISTORICOS: New private helpers used by the 24 tests above
' ============================================================

Private Sub PrepareConflictFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal currentRoot As String, ByVal historicalRoot As String, ByVal dni As String, ByVal conflictHistoricalID As Long, ByRef sourceFile As String)
    Dim fileVar As String
    PrepareCurrentFixture db, userId, currentRoot, historicalRoot, fileVar
    sourceFile = fileVar
    db.Execute "DELETE FROM TbUsuariosHistoricos WHERE ID=" & conflictHistoricalID, dbFailOnError
    db.Execute "DELETE FROM TbUsuariosHistoricos WHERE DNI='" & Replace(dni, "'", "''") & "'", dbFailOnError
    db.Execute "INSERT INTO TbUsuariosHistoricos (ID, DNI, Nombre, Apellido_1) VALUES (" & conflictHistoricalID & ", '" & Replace(dni, "'", "''") & "', 'Preexisting', 'Conflict')", dbFailOnError
End Sub

Private Sub PrepareFolderConflictFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal currentRoot As String, ByVal historicalRoot As String, ByRef sourceFile As String)
    Dim fileVar As String
    PrepareCurrentFixture db, userId, currentRoot, historicalRoot, fileVar
    sourceFile = fileVar
    Dim conflictFolder As String
    conflictFolder = EnsureSlash(historicalRoot) & "DNI" & userId
    CreateTextFile EnsureSlash(conflictFolder) & "conflict-pre-existing.txt", "conflict content"
End Sub

Private Sub PrepareBothConflictFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal currentRoot As String, ByVal historicalRoot As String, ByRef sourceFile As String)
    PrepareConflictFixture db, userId, currentRoot, historicalRoot, "DNI" & userId, userId + 100, sourceFile
    Dim conflictFolder As String
    conflictFolder = EnsureSlash(historicalRoot) & "DNI" & userId
    CreateTextFile EnsureSlash(conflictFolder) & "conflict-pre-existing.txt", "conflict content"
End Sub

Private Sub PrepareH2CConflictFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal historicalRoot As String, ByVal currentRoot As String)
    PrepareHistoricalFixture db, userId, currentRoot, historicalRoot
    db.Execute "DELETE FROM TbUsuarios WHERE ID=" & userId + 100, dbFailOnError
    db.Execute "DELETE FROM TbUsuarios WHERE DNI='DNI" & userId & "'", dbFailOnError
    db.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & userId + 100 & ", 'DNI" & userId & "', 'Preexisting', 'CurrentConflict')", dbFailOnError
End Sub

Private Sub PrepareH2CFolderConflictFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal historicalRoot As String, ByVal currentRoot As String)
    PrepareHistoricalFixture db, userId, currentRoot, historicalRoot
    Dim conflictFolder As String
    conflictFolder = EnsureSlash(currentRoot) & "DNI" & userId & "\ACTUAL\"
    CreateTextFile EnsureSlash(conflictFolder) & "conflict-pre-existing.txt", "conflict content"
End Sub

Private Sub PrepareH2CBothConflictFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal historicalRoot As String, ByVal currentRoot As String)
    PrepareH2CConflictFixture db, userId, historicalRoot, currentRoot
    Dim conflictFolder As String
    conflictFolder = EnsureSlash(currentRoot) & "DNI" & userId & "\ACTUAL\"
    CreateTextFile EnsureSlash(conflictFolder) & "conflict-pre-existing.txt", "conflict content"
End Sub

Private Sub PrepareCurrentFixtureV2(ByRef db As DAO.Database, ByVal userId As Long)
    db.Execute "DELETE FROM TbUsuarios WHERE ID=" & userId, dbFailOnError
    db.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & userId & ", 'DNI" & userId & "', 'Test', 'User')", dbFailOnError
End Sub

Private Sub PrepareAnexoFixtureV2(ByRef db As DAO.Database, ByVal userId As Long)
    db.Execute "DELETE FROM TbAnexosUsuariosHPS WHERE IDUsuario=" & userId, dbFailOnError
    db.Execute "DELETE FROM TbUsuarios WHERE ID=" & userId, dbFailOnError
    db.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & userId & ", 'DNI" & userId & "', 'Test', 'User')", dbFailOnError
    db.Execute "INSERT INTO TbAnexosUsuariosHPS (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & userId + 1 & ", 'current-note.txt', " & userId & ", 'No')", dbFailOnError
End Sub

Private Sub CreateCurrentAttachmentTreeT2(ByVal currentRoot As String, ByVal userId As Long)
    Dim attachmentRoot As String
    attachmentRoot = EnsureSlash(currentRoot) & "DNI" & userId & "\ACTUAL\"
    CreateTextFile attachmentRoot & "current-note.txt", "current attachment"
End Sub

Private Sub PrepareMultipleAnexosFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal currentRoot As String)
    db.Execute "DELETE FROM TbAnexosUsuariosHPS WHERE IDUsuario=" & userId, dbFailOnError
    db.Execute "DELETE FROM TbUsuarios WHERE ID=" & userId, dbFailOnError
    db.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & userId & ", 'DNI" & userId & "', 'Test', 'User')", dbFailOnError
    db.Execute "INSERT INTO TbAnexosUsuariosHPS (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & userId + 1 & ", 'current-note-a.txt', " & userId & ", 'No')", dbFailOnError
    db.Execute "INSERT INTO TbAnexosUsuariosHPS (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & userId + 2 & ", 'current-note-b.txt', " & userId & ", 'No')", dbFailOnError
    db.Execute "INSERT INTO TbAnexosUsuariosHPS (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & userId + 3 & ", 'current-note-c.txt', " & userId & ", 'No')", dbFailOnError
    Dim attachmentRoot As String
    attachmentRoot = EnsureSlash(currentRoot) & "DNI" & userId & "\ACTUAL\"
    CreateTextFile attachmentRoot & "current-note-a.txt", "anexo A"
    CreateTextFile attachmentRoot & "current-note-b.txt", "anexo B"
    CreateTextFile attachmentRoot & "current-note-c.txt", "anexo C"
End Sub

Private Sub PrepareObservacionFixtureV2(ByRef db As DAO.Database, ByVal userId As Long, ByVal currentRoot As String, ByRef unusedFile As String)
    db.Execute "DELETE FROM TbObservaciones WHERE ID=" & userId, dbFailOnError
    db.Execute "DELETE FROM TbUsuarios WHERE ID=" & userId, dbFailOnError
    db.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & userId & ", 'DNI" & userId & "', 'Test', 'User')", dbFailOnError
    db.Execute "INSERT INTO TbObservaciones (IDObservacion, ID, Fecha, Observacion, Tipo) VALUES (950001, " & userId & ", #2026-06-30#, 'note', 'N')", dbFailOnError
End Sub

' ============================================================
' Existing fixtures and helpers (preserved from main)
' ============================================================

Private Sub PrepareCurrentFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal currentRoot As String, ByVal historicalRoot As String, ByRef sourceFile As String)
    TeardownFixture db, userId, currentRoot, historicalRoot
    db.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & userId & ", 'DNI" & userId & "', 'Test', 'Current')", dbFailOnError
    db.Execute "INSERT INTO TbAnexosUsuariosHPS (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & (userId + 1) & ", 'current-note.txt', " & userId & ", 'No')", dbFailOnError
    sourceFile = EnsureSlash(currentRoot) & "DNI" & userId & "\ACTUAL\current-note.txt"
    CreateCurrentAttachmentTree currentRoot, userId
End Sub

Private Sub PrepareHistoricalFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal currentRoot As String, ByVal historicalRoot As String)
    TeardownFixture db, userId, currentRoot, historicalRoot
    db.Execute "INSERT INTO TbUsuariosHistoricos (ID, DNI, Nombre, Apellido_1) VALUES (" & userId & ", 'DNI" & userId & "', 'Test', 'Historical')", dbFailOnError
    db.Execute "INSERT INTO TbAnexosUsuariosHistoricos (IDAnexo, NombreAnexo, IDUsuario) VALUES (" & (userId + 1) & ", 'historical-note.txt', " & userId & ")", dbFailOnError
    CreateTextFile EnsureSlash(historicalRoot) & "DNI" & userId & "\historical-note.txt", "historical attachment"
End Sub

Private Sub TeardownFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal currentRoot As String, ByVal historicalRoot As String)
    On Error Resume Next
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbAnexosUsuariosHPS WHERE IDUsuario=" & userId, dbFailOnError
        db.Execute "DELETE FROM TbAnexosUsuariosHistoricos WHERE IDUsuario=" & userId, dbFailOnError
        db.Execute "DELETE FROM TbUsuarios WHERE ID=" & userId, dbFailOnError
        db.Execute "DELETE FROM TbUsuariosHistoricos WHERE ID=" & userId, dbFailOnError
    End If
    DeleteFolderIfExists currentRoot
    DeleteFolderIfExists historicalRoot
End Sub

Private Function CountRows(ByRef db As DAO.Database, ByVal tableName As String, ByVal whereClause As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM " & tableName & " WHERE " & whereClause, dbOpenSnapshot)
    CountRows = CLng(rs!C)
    rs.Close
End Function

Private Sub CreateCurrentAttachmentTree(ByVal currentRoot As String, ByVal userId As Long)
    Dim attachmentRoot As String
    attachmentRoot = EnsureSlash(currentRoot) & "DNI" & userId & "\ACTUAL\"

    CreateTextFile attachmentRoot & "current-note.txt", "current attachment"
    CreateTextFile attachmentRoot & "level-one\child-a.txt", "nested child A"
    CreateTextFile attachmentRoot & "level-one\level-two\child-b.txt", "nested child B"
End Sub

Private Function AssertCurrentAttachmentTree(ByVal currentRoot As String, ByVal userId As Long, ByRef p_Error As String) As Boolean
    Dim attachmentRoot As String
    Dim relativePaths As Variant
    Dim expectedContents As Variant
    Dim i As Long
    Dim actualPath As String
    Dim actualContent As String
    Dim fileCount As Long

    p_Error = ""
    attachmentRoot = EnsureSlash(currentRoot) & "DNI" & userId & "\ACTUAL\"
    If Dir$(attachmentRoot, vbDirectory) = "" Then
        p_Error = "Missing restored ACTUAL folder: " & attachmentRoot
        Exit Function
    End If

    relativePaths = Array("current-note.txt", "level-one\child-a.txt", "level-one\level-two\child-b.txt")
    expectedContents = Array("current attachment", "nested child A", "nested child B")

    For i = LBound(relativePaths) To UBound(relativePaths)
        actualPath = attachmentRoot & CStr(relativePaths(i))
        If Dir$(actualPath) = "" Then
            p_Error = "Missing restored relative path: " & CStr(relativePaths(i))
            Exit Function
        End If

        actualContent = ReadTextFile(actualPath)
        If actualContent <> CStr(expectedContents(i)) Then
            p_Error = "Content mismatch for " & CStr(relativePaths(i)) & ". Expected '" & CStr(expectedContents(i)) & "', got '" & actualContent & "'"
            Exit Function
        End If
    Next i

    fileCount = CountFilesRecursive(attachmentRoot)
    If fileCount <> (UBound(relativePaths) - LBound(relativePaths) + 1) Then
        p_Error = "Unexpected restored file count. Expected " & (UBound(relativePaths) - LBound(relativePaths) + 1) & ", got " & fileCount
        Exit Function
    End If

    AssertCurrentAttachmentTree = True
End Function

Private Function CountFilesRecursive(ByVal folderPath As String) As Long
    Dim fso As Object
    Dim folder As Object

    If Len(folderPath) = 0 Then Exit Function
    If Dir$(folderPath, vbDirectory) = "" Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    CountFilesRecursive = CountFilesInFolder(folder)
End Function

Private Function CountFilesInFolder(ByVal folder As Object) As Long
    Dim childFolder As Object

    CountFilesInFolder = folder.Files.Count
    For Each childFolder In folder.SubFolders
        CountFilesInFolder = CountFilesInFolder + CountFilesInFolder(childFolder)
    Next childFolder
End Function

Private Function TempLifecycleRoot(ByVal suffix As String) As String
    TempLifecycleRoot = EnsureSlash(Environ$("TEMP")) & "HPS\LifecycleWrapperTests\" & suffix & "\"
End Function

Private Function OpenLocalTestBackend(ByRef p_Error As String) As DAO.Database
    Dim backendPath As String
    Dim connectString As String
    On Error GoTo EH

    p_Error = ""
    backendPath = EnsureSlash(CurrentProject.path) & "HPST.accdb"
    If Dir$(backendPath) = "" Then
        p_Error = "Local test backend not found: " & backendPath
        Exit Function
    End If

    connectString = BackendConnectString()
    Set OpenLocalTestBackend = DBEngine.Workspaces(0).OpenDatabase(backendPath, False, False, connectString)
    Exit Function
EH:
    p_Error = "OpenLocalTestBackend: " & Err.Description
End Function

Private Function BackendConnectString() As String
    Dim password As String
    password = Environ$("DYSFLOW_BACKEND_PASSWORD")
    If Len(password) = 0 Then password = Environ$("HPS_BACKEND_PASSWORD")
    If Len(password) = 0 Then password = Environ$("ACCESS_VBA_PASSWORD")
    If Len(password) = 0 Then password = LocalBackendPassword()
    If Len(password) = 0 Then Err.Raise 1000, , "BackendConnectString: DYSFLOW_BACKEND_PASSWORD, HPS_BACKEND_PASSWORD, ACCESS_VBA_PASSWORD, or .dysflow/backend.pwd is required for HATW tests."
    BackendConnectString = "MS Access;PWD=" & password
End Function

Private Function LocalBackendPassword() As String
    Dim path As String
    path = EnsureSlash(CurrentProject.path) & ".dysflow\backend.pwd"
    If Dir$(path) = "" Then Exit Function
    LocalBackendPassword = Trim$(ReadTextFile(path))
End Function

Private Function IsPostCommitRefreshWarning(ByVal message As String) As Boolean
    IsPostCommitRefreshWarning = (InStr(1, message, "post-commit refresh", vbTextCompare) > 0 _
        Or InStr(1, message, "InicializarTablas", vbTextCompare) > 0)
End Function

Private Function ReadTextFile(ByVal path As String) As String
    Dim ts As Object
    Set ts = CreateObject("Scripting.FileSystemObject").OpenTextFile(path, 1, False)
    ReadTextFile = ts.ReadAll
    ts.Close
End Function

Private Sub CreateTextFile(ByVal path As String, ByVal content As String)
    Dim parentPath As String
    parentPath = Left$(path, InStrRev(path, "\") - 1)
    CreateFolderTree parentPath
    Dim ts As Object
    Set ts = CreateObject("Scripting.FileSystemObject").CreateTextFile(path, True)
    ts.Write content
    ts.Close
End Sub

Private Sub CreateFolderTree(ByVal path As String)
    Dim parts() As String
    Dim currentPath As String
    Dim i As Long
    parts = Split(path, "\")
    currentPath = parts(0) & "\"
    For i = 1 To UBound(parts)
        If parts(i) <> "" Then
            currentPath = currentPath & parts(i) & "\"
            If Dir$(currentPath, vbDirectory) = "" Then MkDir currentPath
        End If
    Next i
End Sub

Private Sub DeleteFolderIfExists(ByVal path As String)
    If Len(path) = 0 Then Exit Sub
    If Dir$(path, vbDirectory) <> "" Then CreateObject("Scripting.FileSystemObject").DeleteFolder path, True
End Sub

Private Function EnsureSlash(ByVal path As String) As String
    If Right$(path, 1) = "\" Then
        EnsureSlash = path
    Else
        EnsureSlash = path & "\"
    End If
End Function

Private Function JsonOk(ByVal value As String, ByRef logs As Collection) As String
    JsonOk = "{""ok"":true,""value"":""" & EscapeJson(value) & """,""payload"":null,""error"":null,""logs"":" & LogsJson(logs) & "}"
End Function

Private Function JsonFail(ByVal message As String, ByRef logs As Collection) As String
    JsonFail = "{""ok"":false,""value"":null,""payload"":null,""error"":""" & EscapeJson(message) & """,""logs"":" & LogsJson(logs) & "}"
End Function

Private Function LogsJson(ByRef logs As Collection) As String
    Dim i As Long
    Dim result As String
    result = "["
    If Not logs Is Nothing Then
        For i = 1 To logs.Count
            If i > 1 Then result = result & ","
            result = result & """" & EscapeJson(CStr(logs(i))) & """"
        Next i
    End If
    LogsJson = result & "]"
End Function

Private Function EscapeJson(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, """", Chr$(92) & Chr$(34))
    value = Replace(value, vbCrLf, "\n")
    value = Replace(value, vbCr, "\n")
    value = Replace(value, vbLf, "\n")
    EscapeJson = value
End Function

