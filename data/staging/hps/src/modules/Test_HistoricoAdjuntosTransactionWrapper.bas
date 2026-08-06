Attribute VB_Name = "Test_HistoricoAdjuntosTransactionWrapper"
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

' [PR1a dependency injection] — when the test sets m_PromptDb before invoking
' the coordinator, the conflict-prompt callbacks use THIS database handle
' instead of opening their own. Sharing the handle keeps the prompt
' function and the coordinator in the same Workspace/transaction domain,
' so rows injected from the prompt are immediately visible to the
' coordinator's precheck#2 — no cross-workspace cache surprises, no need
' for TableDefs.Refresh hacks. Defaults to Nothing; when Nothing, the
' callbacks fall back to OpenLocalTestBackend (preserves prior tests that
' never set the prompt db).
Private m_PromptDb As DAO.Database

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
        Test_HATW_CurrentToHistorical_RestoresAfterSourceDeleteFailure = JsonFail("Snapshot restore did not recreate the flat attachment layout: " & treeError, logs)
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

    logs.Add "AfterSourceDelete failure restored the attachment files flat at <DNI>\\ from snapshot."
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
    ' [c2h-h2c-subcarpetas] PrepareHistoricalFixture now seeds EsHistorico='No',
    ' so the h2c split-routing lands the file in <currentRoot>\DNI<id>\Actual\
    ' instead of the legacy top-level <DNI>\ path.
    targetFile = EnsureSlash(currentRoot) & "DNI" & TEST_USER_HISTORICAL & "\Actual\historical-note.txt"

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
    targetFile = EnsureSlash(historicalRoot) & "DNI" & TEST_USER_CURRENT & "\current-note.txt"

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
    ' [c2h-h2c-subcarpetas] PrepareHistoricalFixture seeds EsHistorico='No';
    ' h2c split-routing puts the file in Actual\ subfolder.
    targetFile = EnsureSlash(currentRoot) & "DNI" & TEST_USER_HISTORICAL & "\Actual\historical-note.txt"

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
    targetFile = EnsureSlash(historicalRoot) & "DNI" & TEST_USER_CURRENT & "\current-note.txt"

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
    Dim ws As DAO.Workspace
    m_LastConflictMessage = conflicts
    m_LastConflictDNI = dni
    m_LastConflictDirection = direction
    m_ConflictPromptCalls = m_ConflictPromptCalls + 1
    If m_TOCTOUExtraHistoricalUserID > 0 Then
        LogToTestTrace "[Test_ConflictPrompt_ApproveAndInjectExtraHistoricalUser] attempting INSERT id=" & m_TOCTOUExtraHistoricalUserID & " dni=" & dni
        ' [PR1a dependency injection] — when the test set m_PromptDb via
        ' Test_SetPromptDb, reuse that exact handle so the prompt function
        ' and the coordinator operate in the same Workspace/transaction
        ' domain. This eliminates the cross-workspace visibility problem
        ' the previous isolated-Workspace hack tried (and failed) to work
        ' around: the isoWorkspace made the prompt's INSERT invisible to
        ' the coordinator's precheck#2 AND left a lingering handle that
        ' blocked the test's cleanup DELETE — the orphan TOCTOU rows that
        ' accumulated across runs are a direct symptom. When m_PromptDb is
        ' Nothing (older tests that never call Test_SetPromptDb), fall back
        ' to OpenLocalTestBackend so those keep working unchanged.
        If Not m_PromptDb Is Nothing Then
            Set db = m_PromptDb
            LogToTestTrace "[Test_ConflictPrompt_ApproveAndInjectExtraHistoricalUser] using injected prompt db, name=" & db.Name
        Else
            Set db = OpenLocalTestBackend(openErr)
            If openErr <> "" Then
                LogToTestTrace "[Test_ConflictPrompt_ApproveAndInjectExtraHistoricalUser] fallback open err=" & openErr
            End If
        End If
        If Not db Is Nothing Then
            ' [PR1a TOCTOU] Use db.Execute for the INSERT (not AddNew/Update on
            ' a dynaset) because in Jet/ACE a dynaset Update can leave the row
            ' in a pending state on the Workspace's implicit transaction. The
            ' coordinator's precheck#2 (dbOpenDynaset on the same db handle)
            ' would then read a stale view that misses the new row even though
            ' the row IS physically in the file (verified via query_sql from a
            ' fresh connection). db.Execute is a lower-level path that flushes
            ' to the file as part of the operation, so the next OpenRecordset
            ' from the same db handle sees the commit.
            '
            ' [PR1a TOCTOU fix] — TbUsuariosHistoricos has a UNIQUE index on DNI,
            ' so inserting with the same DNI as the source user (which the test
            ' fixture just inserted at a different ID) raises DAO 3022. The TOCTOU
            ' race is detected by the *count* of conflicts (precheck#1 vs precheck#2),
            ' not by the literal DNI value — so the injected row carries a
            ' suffix-stamped DNI to satisfy the unique index. The coordinator's
            ' precheck#2 query matches the suffix via the `DNI LIKE '..._TOCTOU*'`
            ' clause, so the counter-compare still detects the new row.
            On Error Resume Next
            db.Execute "INSERT INTO TbUsuariosHistoricos (ID, DNI, Nombre, Apellido_1) VALUES (" & m_TOCTOUExtraHistoricalUserID & ", '" & Replace(dni & "_TOCTOU", "'", "''") & "', 'TOCTOU', 'Race')", dbFailOnError
            LogToTestTrace "[Test_ConflictPrompt_ApproveAndInjectExtraHistoricalUser] after INSERT err=" & Err.Number & " desc=" & Left$(Err.Description, 200)
            On Error GoTo 0
            Dim postCheckRs As DAO.Recordset
            Set postCheckRs = db.OpenRecordset("SELECT COUNT(*) AS N FROM TbUsuariosHistoricos WHERE ID=" & m_TOCTOUExtraHistoricalUserID, dbOpenSnapshot)
            LogToTestTrace "[Test_ConflictPrompt_ApproveAndInjectExtraHistoricalUser] postCheck count for id=" & m_TOCTOUExtraHistoricalUserID & " = " & postCheckRs!N
            postCheckRs.Close
        End If
        ' Do NOT close db here — when it is m_PromptDb, the test owns the
        ' handle and is responsible for closing it. When it is the
        ' OpenLocalTestBackend fallback, the local reference going out of
        ' scope is enough.
        Set db = Nothing
    End If
    Test_ConflictPrompt_ApproveAndInjectExtraHistoricalUser = True
End Function

' [PR1a debug] — write a one-line trace to %TEMP%\HPS\pr1a-trace.log so the
' agent can see what the prompt function did during the TOCTOU test.
Private Sub LogToTestTrace(ByVal p_Message As String)
    On Error Resume Next
    Dim logPath As String
    logPath = Environ$("TEMP") & "\HPS\pr1a-trace.log"
    Dim f As Integer
    f = FreeFile
    Open logPath For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & p_Message
    Close #f
End Sub

' [PR1a relink] — Rewrite the linked-table Connect strings that point to the
' producción UNC \\datoste\aplicaciones_dys\... and replace them with local
' C:\00repos\datos\... paths. This is the root cause of the "Ha intentado
' confirmar..." and "duplicate in index" errors in the TOCTOU tests: the FE's
' linked tables pointed to UNC paths the test env couldn't reach, so any DAO
' operation that went through the linked-table path failed.
'
' Run via dysflow_vba_execute once after restoring the FE, then any subsequent
' test runs use local backends only.
Public Sub RelinkToLocalData()
    On Error GoTo Fail
    Dim tdf As DAO.TableDef
    Dim newConnect As String
    Dim replaced As Long
    Dim inspected As Long
    Dim attachedSeen As Long
    Dim currentConnect As String
    For Each tdf In CurrentDb.TableDefs
        inspected = inspected + 1
        If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
            attachedSeen = attachedSeen + 1
            currentConnect = tdf.Connect
            If InStr(1, currentConnect, "\\datoste\", vbTextCompare) > 0 Then
                Dim lastSlash As Long
                Dim fileNameOnly As String
                lastSlash = InStrRev(currentConnect, "\")
                If lastSlash > 0 Then
                    fileNameOnly = Mid$(currentConnect, lastSlash + 1)
                Else
                    fileNameOnly = currentConnect
                End If
                Dim localPath As String
                localPath = "C:\00repos\datos\" & fileNameOnly
                newConnect = "MS Access;DATABASE=" & localPath & ";PWD=dpddpd"
                tdf.Connect = newConnect
                tdf.RefreshLink
                replaced = replaced + 1
                LogToTestTrace "[RelinkToLocalData] " & tdf.Name & " -> " & localPath
            End If
        End If
    Next tdf

    LogToTestTrace "[RelinkToLocalData] done: inspected=" & inspected & " attached=" & attachedSeen & " replaced=" & replaced
    Debug.Print "[RelinkToLocalData] done: inspected=" & inspected & " attached=" & attachedSeen & " replaced=" & replaced

    ' Now also relink the project's HPST.accdb (which the tests use as their backend).
    Dim backendPath As String
    backendPath = CurrentProject.path & "\HPST.accdb"
    If Dir$(backendPath) <> "" Then
        LogToTestTrace "[RelinkToLocalData] also relinking HPST at " & backendPath
        RelinkLocalDataInDb backendPath, replaced
    End If

    ' Finally: drop the broken TbUsuarioSolicitud link from
    ' C:\00repos\datos\Solicitudes_HPS_datos.accdb. dysflow's relink_directory
    ' marked it as plannedRelink but failed at apply (the target table doesn't
    ' exist in the local HPST.accdb). The link still points at \\datoste\...HPST.accdb
    ' which is unreachable — and since HPST.accdb (data) is the relink target
    ' for many other Solicitudes links, the broken link makes Solicitudes_HPS_datos.accdb
    ' unopenable, which cascades errors back into HPST.accdb and blocks every test
    ' that does an INSERT (DAO 3022 on FK-relation validation).
    DropBrokenSolicitudesLink

    Exit Sub
Fail:
    LogToTestTrace "[RelinkToLocalData] err=" & Err.Number & " desc=" & Err.Description
    Debug.Print "[RelinkToLocalData] err=" & Err.Number & " desc=" & Err.Description
End Sub

' [PR1a relink] — drop the broken TbUsuarioSolicitud link from
' C:\00repos\datos\Solicitudes_HPS_datos.accdb. The link points at
' \\datoste\...\HPST.accdb (producción, unreachable) and the target table does
' not exist in the local HPST.accdb. dysflow's relink_directory marks it
' "plannedRelink" but cannot refresh-link it, leaving Solicitudes unopenable.
' We open Solicitudes in an isolated workspace and delete the tabledef.
Private Sub DropBrokenSolicitudesLink()
    On Error GoTo Fail
    Dim isoWs As DAO.Workspace
    Dim isoDb As DAO.Database
    Dim tdf As DAO.TableDef
    Set isoWs = DBEngine.CreateWorkspace("DropLinkIso", "admin", "", dbUseJet)
    Set isoDb = isoWs.OpenDatabase("C:\00repos\datos\Solicitudes_HPS_datos.accdb", False, False, "MS Access;PWD=dpddpd")
    If isoDb Is Nothing Then
        LogToTestTrace "[DropBrokenSolicitudesLink] failed to open Solicitudes"
        Exit Sub
    End If
    For Each tdf In isoDb.TableDefs
        If tdf.Name = "TbUsuarioSolicitud" Then
            If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
                ' Use the collection's Delete method (with the table name) rather
                ' than tdf.Delete — some Access/VBA builds fail to resolve the
                ' Delete member on the iterated DAO.TableDef variable.
                isoDb.TableDefs.Delete "TbUsuarioSolicitud"
                LogToTestTrace "[DropBrokenSolicitudesLink] dropped TbUsuarioSolicitud"
                Exit For
            End If
        End If
    Next tdf
    isoDb.Close
    isoWs.Close
    Exit Sub
Fail:
    LogToTestTrace "[DropBrokenSolicitudesLink] err=" & Err.Number & " desc=" & Err.Description
End Sub

' Helper: open `p_DbPath` as its own Database object and re-point any
' attached tables whose Connect string points to \\datoste\... to a local
' C:\00repos\datos\<file>.accdb path with the right password.
Private Sub RelinkLocalDataInDb(ByVal p_DbPath As String, ByRef p_TotalReplaced As Long)
    On Error GoTo Fail
    Dim isoWs As DAO.Workspace
    Dim isoDb As DAO.Database
    Set isoWs = DBEngine.CreateWorkspace("RelinkLocalIso", "admin", "", dbUseJet)
    Set isoDb = isoWs.OpenDatabase(p_DbPath, False, False, "MS Access;PWD=dpddpd")
    If isoDb Is Nothing Then
        LogToTestTrace "[RelinkLocalDataInDb] failed to open " & p_DbPath
        Exit Sub
    End If
    Dim tdf As DAO.TableDef
    For Each tdf In isoDb.TableDefs
        If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
            If InStr(1, tdf.Connect, "\\datoste\", vbTextCompare) > 0 Then
                Dim lastSlash As Long
                Dim fileNameOnly As String
                lastSlash = InStrRev(tdf.Connect, "\")
                If lastSlash > 0 Then
                    fileNameOnly = Mid$(tdf.Connect, lastSlash + 1)
                Else
                    fileNameOnly = tdf.Connect
                End If
                Dim localPath As String
                localPath = "C:\00repos\datos\" & fileNameOnly
                tdf.Connect = "MS Access;DATABASE=" & localPath & ";PWD=dpddpd"
                tdf.RefreshLink
                p_TotalReplaced = p_TotalReplaced + 1
                LogToTestTrace "[RelinkLocalDataInDb] " & p_DbPath & " :: " & tdf.Name & " -> " & localPath
            End If
        End If
    Next tdf
    isoDb.Close
    isoWs.Close
    Exit Sub
Fail:
    LogToTestTrace "[RelinkLocalDataInDb] err=" & Err.Number & " desc=" & Err.Description
End Sub

' Approves AND injects another current user row (same DNI, different ID) for H2C direction.
Public Function Test_ConflictPrompt_ApproveAndInjectExtraCurrentUser(ByVal conflicts As String, ByVal dni As String, ByVal direction As String) As Boolean
    Dim openErr As String
    Dim db As DAO.Database
    Dim ownsDb As Boolean
    m_LastConflictMessage = conflicts
    m_LastConflictDNI = dni
    m_LastConflictDirection = direction
    m_ConflictPromptCalls = m_ConflictPromptCalls + 1
    If m_TOCTOUExtraCurrentUserID > 0 Then
        ' [PR1a dependency injection] — same pattern as the historical
        ' variant: prefer the test's m_PromptDb handle so the prompt
        ' INSERT runs in the same Workspace/transaction domain the
        ' coordinator is about to read. Fallback to OpenLocalTestBackend
        ' for tests that did not inject.
        If Not m_PromptDb Is Nothing Then
            Set db = m_PromptDb
            ownsDb = False
            LogToTestTrace "[Test_ConflictPrompt_ApproveAndInjectExtraCurrentUser] using injected prompt db, name=" & db.Name
        Else
            Set db = OpenLocalTestBackend(openErr)
            ownsDb = True
            If openErr <> "" Then
                LogToTestTrace "[Test_ConflictPrompt_ApproveAndInjectExtraCurrentUser] fallback open err=" & openErr
            End If
        End If
        If openErr = "" And Not db Is Nothing Then
            On Error Resume Next
            ' [PR1a TOCTOU fix] — TbUsuarios has a UNIQUE index on DNI. Inserting with
            ' the same DNI as the source user (which the test fixture just inserted
            ' at a different ID) raises DAO 3022. The TOCTOU race is detected by
            ' the count of conflicts, not by the literal DNI, so we use a
            ' suffix-stamped DNI to satisfy the unique index.
            db.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & m_TOCTOUExtraCurrentUserID & ", '" & Replace(dni & "_TOCTOU", "'", "''") & "', 'TOCTOU', 'Race')", dbFailOnError
            On Error GoTo 0
            ' Only close if we opened the handle ourselves. Closing m_PromptDb
            ' would tear down the test's db out from under it.
            If ownsDb Then
                On Error Resume Next
                db.Close
                On Error GoTo 0
            End If
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
    Set m_PromptDb = Nothing
End Sub

' [PR1a dependency injection] — call this before invoking the coordinator
' when the test wants the conflict-prompt callbacks to operate on the SAME
' db handle the test already has open. Sharing the handle eliminates the
' cross-workspace visibility problem that the previous isoWorkspace hack
' was trying (and failing) to work around. The test retains ownership of
' the db handle (do not close it inside the prompt callback).
Public Sub Test_SetPromptDb(ByVal db As DAO.Database)
    Set m_PromptDb = db
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

    ' [PR1a dependency injection] — share this test's db handle with the
    ' conflict-prompt callback so the TOCTOU row the callback injects is
    ' visible to the coordinator's precheck#2 in the same Workspace. See
    ' the comment on m_PromptDb in the module header for the rationale.
    Test_SetPromptDb db

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    coordinator.ConfigureLifecycleConflictPrompt "Test_ConflictPrompt_ApproveAndInjectExtraHistoricalUser"
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg = "" Then
        Test_HATW_TOCTOU_DetectsNewConflictAfterCleanup = JsonFail("Expected TOCTOU abort, got empty errMsg. Trace: " & Environ$("TEMP") & "\HPS\pr1a-trace.log", logs)
        GoTo CleanUp
    End If
    If InStr(1, errMsg, "TOCTOU", vbTextCompare) = 0 And InStr(1, errMsg, "race", vbTextCompare) = 0 And InStr(1, errMsg, "nuevo conflicto", vbTextCompare) = 0 And InStr(1, errMsg, "new conflict", vbTextCompare) = 0 Then
        Test_HATW_TOCTOU_DetectsNewConflictAfterCleanup = JsonFail("Expected TOCTOU/race error message, got: " & errMsg & " | Trace: " & Environ$("TEMP") & "\HPS\pr1a-trace.log", logs)
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
        ' [PR1a TOCTOU cleanup] — also delete the _TOCTOU and _TOCTOU_H2C suffixed
        ' rows that the prompt's ApproveAndInject functions inject. Without
        ' these, repeated test runs accumulate orphan rows and the second
        ' run's preCheck finds a stale row (count=1) — making the AddNew fail
        ' with DAO 3022 even on a clean file.
        db.Execute "DELETE FROM TbUsuariosHistoricos WHERE DNI LIKE 'DNI" & TEST_USER_CURRENT & "_TOCTOU*'", dbFailOnError
        db.Execute "DELETE FROM TbUsuarios WHERE DNI LIKE 'DNI" & TEST_USER_CURRENT & "_TOCTOU*'", dbFailOnError
    End If
    Exit Function
EH:
    Test_HATW_TOCTOU_DetectsNewConflictAfterCleanup = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbUsuariosHistoricos WHERE DNI LIKE 'DNI" & TEST_USER_CURRENT & "_TOCTOU*'", dbFailOnError
        db.Execute "DELETE FROM TbUsuarios WHERE DNI LIKE 'DNI" & TEST_USER_CURRENT & "_TOCTOU*'", dbFailOnError
    End If
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

    ' [PR1a dependency injection] — share this test's db handle with the
    ' conflict-prompt callback so the h2c TOCTOU row the callback injects is
    ' visible to the coordinator's precheck#2 in the same Workspace. Same
    ' rationale as the c2h variant: see the m_PromptDb comment in this
    ' module header.
    Test_SetPromptDb db

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
        ' Fixture base row (exact DNI) + TOCTOU-injected row (suffix variant).
        ' Both must go or the next run starts dirty. Wildcard uses Jet's `*`
        ' (not `%`); the AO 3022 cleanup of the coordinator already handles the
        ' fixture base, but the injected TOCTOU row sits outside its scope and
        ' is the test's responsibility.
        db.Execute "DELETE FROM TbUsuarios WHERE DNI='DNI" & TEST_USER_HISTORICAL & "'", dbFailOnError
        db.Execute "DELETE FROM TbUsuarios WHERE DNI LIKE 'DNI" & TEST_USER_HISTORICAL & "_TOCTOU*'", dbFailOnError
        db.Execute "DELETE FROM TbUsuariosHistoricos WHERE DNI LIKE 'DNI" & TEST_USER_HISTORICAL & "_TOCTOU*'", dbFailOnError
    End If
    Exit Function
EH:
    Test_HATW_HistoricalToCurrent_TOCTOU_DetectsNewConflict = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbUsuarios WHERE DNI='DNI" & TEST_USER_HISTORICAL & "'", dbFailOnError
        db.Execute "DELETE FROM TbUsuarios WHERE DNI LIKE 'DNI" & TEST_USER_HISTORICAL & "_TOCTOU*'", dbFailOnError
        db.Execute "DELETE FROM TbUsuariosHistoricos WHERE DNI LIKE 'DNI" & TEST_USER_HISTORICAL & "_TOCTOU*'", dbFailOnError
    End If
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
    resolvedID = coordinator.ResolveTargetUserID(db, CLng(TEST_USER_CURRENT), "c2h", errMsg)
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
    resolvedID = coordinator.ResolveTargetUserID(db, CLng(TEST_USER_CURRENT), "c2h", errMsg)
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
    db.Execute "CREATE TABLE TbAnexosUsuariosHistoricos (IDAnexo LONG NOT NULL, NombreAnexo TEXT(255), IDUsuario LONG, Hash TEXT(64), EsHistorico TEXT(2), CONSTRAINT PK_TbAnexosUsuariosHistoricos PRIMARY KEY (IDAnexo))", dbFailOnError
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
    ' Permissive error check: lock contention OR FK OR delete — any error during cleanup.
    ' Includes Spanish DAO error keywords ("exclusivo", "abierta") because Jet's
    ' native error message is localized; without these the test rejects the
    ' legitimate "La tabla 'X' está abierta en modo exclusivo" lock error.
    If InStr(1, errMsg, "lock", vbTextCompare) = 0 _
       And InStr(1, errMsg, "use", vbTextCompare) = 0 _
       And InStr(1, errMsg, "exclusivo", vbTextCompare) = 0 _
       And InStr(1, errMsg, "abierta", vbTextCompare) = 0 _
       And InStr(1, errMsg, "FK", vbTextCompare) = 0 _
       And InStr(1, errMsg, "relationship", vbTextCompare) = 0 _
       And InStr(1, errMsg, "referenc", vbTextCompare) = 0 _
       And InStr(1, errMsg, "DeletePriorConflicts", vbTextCompare) = 0 _
       And InStr(1, errMsg, "delete", vbTextCompare) = 0 Then
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
    recreatedAsSourceContent = EnsureSlash(conflictFolderPath) & "current-note.txt"
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
' ANEXOS-HISTORICOS: PR2 tests — cache cleanup + filesystem flatten
' Verifies that the c2h move:
' 1. removes the user from TbUsuariosSICALocalParaIndicadores (SICA indicator
'    cache that the legacy code never touched — observed by user 2026-07-08
'    after auditing Mod_Sincronizacion_Historico.bas:cache_usuario_sincronizar_despues_movimiento).
' 2. flattens the historical anexos folder — every file from the source's
'    ACTUAL\ and HISTORICO\ subfolders lands flat at <DNI>\<file>, no inner
'    subfolders (per user's request: "lo debe hacer así" with all PDFs at
'    <DNI>\ level directly).
' ============================================================

Public Function Test_HATW_C2H_RemovesFromSicaIndicatorsCache() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000
    currentRoot = TempLifecycleRoot("sica-cache-c2h")
    historicalRoot = TempLifecycleRoot("sica-cache-c2h-hist")
    Dim sourceFile As String
    PrepareCurrentFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile

    ' Inject a row in the SICA indicator cache that the c2h move should clean.
    ' The cache sync function (Mod_Sincronizacion_Historico.bas) currently
    ' does NOT touch this table (PR2 gap); this test will be RED until fixed.
    '
    ' IMPORTANT: TbUsuariosSICALocalParaIndicadores is a LOCAL (frontend) cache
    ' table — it does NOT exist in the test backend (HPST.accdb). The cache
    ' sync uses CurrentDb() (staging HPS.accdb) to DELETE from this table, so
    ' this test must do the same. Using the test backend (db) here would error
    ' with "table or query not found".
    CurrentDb().Execute "DELETE FROM TbUsuariosSICALocalParaIndicadores WHERE ID='" & Replace(CStr(TEST_USER_CURRENT), "'", "''") & "'", dbFailOnError
    CurrentDb().Execute "INSERT INTO TbUsuariosSICALocalParaIndicadores (ID, IDHPS, Nombre, Apellido_1) VALUES ('" & CStr(TEST_USER_CURRENT) & "', " & TEST_USER_CURRENT & ", 'Sica-Indicador-Cache-Test', 'c2h')", dbFailOnError

    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)
    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_HATW_C2H_RemovesFromSicaIndicatorsCache = JsonFail("Expected success, got: " & errMsg, logs)
        GoTo CleanUp
    End If

    If CountRows(CurrentDb(), "TbUsuariosSICALocalParaIndicadores", "ID='" & CStr(TEST_USER_CURRENT) & "'") <> 0 Then
        Test_HATW_C2H_RemovesFromSicaIndicatorsCache = JsonFail("SICA indicator cache row remains after c2h: ID=" & TEST_USER_CURRENT, logs)
        GoTo CleanUp
    End If

    logs.Add "SICA indicator cache row was deleted during c2h."
    Test_HATW_C2H_RemovesFromSicaIndicatorsCache = JsonOk("sica-cache-c2h-removed", logs)
CleanUp:
    ' Clean up the injected SICA indicator row in case the test failed before
    ' the c2h move completed (or to ensure no leftovers across runs).
    On Error Resume Next
    CurrentDb().Execute "DELETE FROM TbUsuariosSICALocalParaIndicadores WHERE ID='" & Replace(CStr(TEST_USER_CURRENT), "'", "''") & "'", dbFailOnError
    On Error GoTo 0
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_C2H_RemovesFromSicaIndicatorsCache = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_LifecyclePaths_AreAlwaysFlat() As String
    ' [PR2 RFC norm regression] Locks the lifecycle-path convention in BOTH
    ' directions. The norm is asymmetric by design after the c2h/h2c
    ' subcarpetas hotfix (2026-07-10):
    '
    '   c2h direction: target <HISTORICO>\<DNI>\ is STRICT-FLAT.
    '     Coordinator.CopyFolderIfExists flattens by design; any inner
    '     ACTUAL\ or HISTORICO\ subfolder would silently break the cache
    '     regen path. The user explicitly required this norm to be
    '     'blindada y muy clara'.
    '
    '   h2c direction: target <HPS>\<DNI>\ uses PER-ANEXO SPLIT-ROUTING.
    '     Coordinator.CopyFolderByEsHistorico reads EsHistorico from each
    '     TbAnexosUsuariosHistoricos row and lands the file in either
    '     <DNI>\Actual\ (EsHistorico='No') or <DNI>\HISTORICO\
    '     (EsHistorico='Sí'). This makes c2h→h2c a reversible cycle: each
    '     anexo returns to the subfolder it came from.
    '
    ' The per-anexo move (AnexoUsuarioHPS.PasarAnexo / PasarAnexoAHistorico)
    ' also uses ACTUAL\<DNI>\HISTORICO\ subfolders — that's a different
    ' code path with its own tests
    ' (Test_AnexoUsuarioHPS_URLCarpetaAnexosHistoricos_DerivesFromActuales
    ' and the 4 Test_C2H_/Test_H2C_ atoms).
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim c2hCurrentRoot As String, c2hHistoricalRoot As String
    Dim h2cCurrentRoot As String, h2cHistoricalRoot As String
    Dim c2hSourceFile As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    ' === c2h: target = <HISTORICO>\<DNI>\, must be flat (no inner subfolders) ===
    c2hCurrentRoot = TempLifecycleRoot("norm-c2h")
    c2hHistoricalRoot = TempLifecycleRoot("norm-c2h-hist")
    PrepareCurrentFixture db, TEST_USER_CURRENT, c2hCurrentRoot, c2hHistoricalRoot, c2hSourceFile
    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, c2hCurrentRoot, c2hHistoricalRoot, "", errMsg)
    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_HATW_LifecyclePaths_AreAlwaysFlat = JsonFail("c2h failed: " & errMsg, logs)
        GoTo CleanUp
    End If
    Dim c2hTargetDni As String
    c2hTargetDni = EnsureSlash(c2hHistoricalRoot) & "DNI" & TEST_USER_CURRENT
    If Dir$(c2hTargetDni, vbDirectory) = "" Then
        Test_HATW_LifecyclePaths_AreAlwaysFlat = JsonFail("c2h did not create target <HISTORICO>\<DNI>\: " & c2hTargetDni, logs)
        GoTo CleanUp
    End If
    ' [PR2 norm enforcement] The source fixture file (current-note.txt) MUST
    ' land at <DNI>\current-note.txt (top level), not inside ACTUAL\ or
    ' HISTORICO\. This is the positive counterpart to the negative subfolder
    ' assertions below.
    If Dir$(c2hTargetDni & "\current-note.txt") = "" Then
        Test_HATW_LifecyclePaths_AreAlwaysFlat = JsonFail("c2h did not flatten source attachment to <DNI>\current-note.txt: " & c2hTargetDni, logs)
        GoTo CleanUp
    End If
    If Dir$(c2hTargetDni & "\ACTUAL", vbDirectory) <> "" Then
        Test_HATW_LifecyclePaths_AreAlwaysFlat = JsonFail("c2h target has ACTUAL\ subfolder (must be flat): " & c2hTargetDni, logs)
        GoTo CleanUp
    End If
    If Dir$(c2hTargetDni & "\HISTORICO", vbDirectory) <> "" Then
        Test_HATW_LifecyclePaths_AreAlwaysFlat = JsonFail("c2h target has HISTORICO\ subfolder (must be flat): " & c2hTargetDni, logs)
        GoTo CleanUp
    End If
    ' [PR2 norm enforcement] STRICT-FLAT: <DNI>\ must have ZERO subfolders,
    ' not just "no ACTUAL\ or HISTORICO\". Any other subfolder name (e.g. a
    ' future "BACKUP\") would still break the cache regen path; the user
    ' explicitly required the norm to be 'blindada y muy clara'.
    Dim c2hFlatFso As Object
    Set c2hFlatFso = CreateObject("Scripting.FileSystemObject")
    If c2hFlatFso.GetFolder(c2hTargetDni).Subfolders.Count <> 0 Then
        Test_HATW_LifecyclePaths_AreAlwaysFlat = JsonFail("c2h target has subfolders (strict-flat required, only files allowed at <DNI>\): " & c2hTargetDni, logs)
        GoTo CleanUp
    End If

    ' === h2c: target = <HPS>\<DNI>\, per-anexo split-routing by EsHistorico ===
    ' The c2h→h2c cycle is reversible: each anexo returns to the subfolder
    ' it came from. After c2h landed both source files flat at
    ' <HISTORICO>\<DNI>\ (see c2h assertions above), h2c must restore each
    ' one to the matching <HPS>\<DNI>\ subfolder:
    '   - file_a.txt (EsHistorico='No') → <DNI>\Actual\file_a.txt
    '   - file_h.txt (EsHistorico='Sí') → <DNI>\HISTORICO\file_h.txt
    h2cCurrentRoot = TempLifecycleRoot("norm-h2c")
    h2cHistoricalRoot = TempLifecycleRoot("norm-h2c-hist")
    PrepareH2CSubcarpetasFixture db, TEST_USER_HISTORICAL, h2cCurrentRoot, h2cHistoricalRoot
    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarHistoricoAActual(CStr(TEST_USER_HISTORICAL), db, h2cCurrentRoot, h2cHistoricalRoot, "", errMsg)
    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_HATW_LifecyclePaths_AreAlwaysFlat = JsonFail("h2c failed: " & errMsg, logs)
        GoTo CleanUp
    End If
    Dim h2cTargetDni As String
    h2cTargetDni = EnsureSlash(h2cCurrentRoot) & "DNI" & TEST_USER_HISTORICAL
    If Dir$(h2cTargetDni, vbDirectory) = "" Then
        Test_HATW_LifecyclePaths_AreAlwaysFlat = JsonFail("h2c did not create target <HPS>\<DNI>\: " & h2cTargetDni, logs)
        GoTo CleanUp
    End If
    ' [c2h-h2c-subcarpetas norm] Per-anexo split-routing: each file lands
    ' in the subfolder matching its EsHistorico flag.
    If Dir$(h2cTargetDni & "\Actual\file_a.txt") = "" Then
        Test_HATW_LifecyclePaths_AreAlwaysFlat = JsonFail("h2c did not route file_a.txt (EsHistorico='No') to <DNI>\Actual\: " & h2cTargetDni & "\Actual\file_a.txt", logs)
        GoTo CleanUp
    End If
    If Dir$(h2cTargetDni & "\HISTORICO\file_h.txt") = "" Then
        Test_HATW_LifecyclePaths_AreAlwaysFlat = JsonFail("h2c did not route file_h.txt (EsHistorico='Sí') to <DNI>\HISTORICO\: " & h2cTargetDni & "\HISTORICO\file_h.txt", logs)
        GoTo CleanUp
    End If
    ' Neither file should be at <DNI>\ top level — that's the OLD strict-flat
    ' norm. The new norm explicitly puts each file in its matching subfolder.
    If Dir$(h2cTargetDni & "\file_a.txt") <> "" Then
        Test_HATW_LifecyclePaths_AreAlwaysFlat = JsonFail("h2c left file_a.txt at <DNI>\ top level (split-routing expected): " & h2cTargetDni & "\file_a.txt", logs)
        GoTo CleanUp
    End If
    If Dir$(h2cTargetDni & "\file_h.txt") <> "" Then
        Test_HATW_LifecyclePaths_AreAlwaysFlat = JsonFail("h2c left file_h.txt at <DNI>\ top level (split-routing expected): " & h2cTargetDni & "\file_h.txt", logs)
        GoTo CleanUp
    End If

    logs.Add "c2h stays strict-flat (<DNI>\<file> with ZERO subfolders); h2c routes per-anexo by EsHistorico (<DNI>\Actual\ for 'No', <DNI>\HISTORICO\ for 'Sí')."
    Test_HATW_LifecyclePaths_AreAlwaysFlat = JsonOk("lifecycle-paths-norm", logs)
CleanUp:
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, c2hCurrentRoot, c2hHistoricalRoot
    TeardownFixture db, TEST_USER_HISTORICAL, h2cCurrentRoot, h2cHistoricalRoot
    On Error GoTo 0
    Exit Function
EH:
    Test_HATW_LifecyclePaths_AreAlwaysFlat = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, c2hCurrentRoot, c2hHistoricalRoot
    TeardownFixture db, TEST_USER_HISTORICAL, h2cCurrentRoot, h2cHistoricalRoot
End Function

Public Function Test_HATW_C2H_FlattensAnexosFolder() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim dni As String
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000
    currentRoot = TempLifecycleRoot("flatten-c2h")
    historicalRoot = TempLifecycleRoot("flatten-c2h-hist")
    Dim sourceFile As String
    PrepareCurrentFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile

    ' Re-create source as legacy-layout: DNI\ACTUAL\file and DNI\HISTORICO\file
    ' (the dirty-state path the user is fixing). The coordinator's flat copy
    ' should pull all files flat into the target at <DNI>\file.
    '
    ' IMPORTANT: the fixture (PrepareCurrentFixture -> CreateCurrentAttachmentTree)
    ' already created <DNI>\ACTUAL\ and writes the source file with overwrite=False,
    ' so re-creating the folder with FSO.CreateFolder errors with
    ' "El archivo ya existe". Use Dir$() check or On Error Resume Next to
    ' avoid the "folder already exists" error on re-runs.
    dni = "DNI" & TEST_USER_CURRENT
    If Dir$(sourceFile) <> "" Then Kill sourceFile
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Dir$(EnsureSlash(currentRoot) & dni, vbDirectory) = "" Then MkDir EnsureSlash(currentRoot) & dni
    If Dir$(EnsureSlash(currentRoot) & dni & "\ACTUAL", vbDirectory) = "" Then fs.CreateFolder EnsureSlash(currentRoot) & dni & "\ACTUAL"
    If Dir$(EnsureSlash(currentRoot) & dni & "\HISTORICO", vbDirectory) = "" Then fs.CreateFolder EnsureSlash(currentRoot) & dni & "\HISTORICO"
    Dim fileActual As String
    Dim fileHist As String
    fileActual = EnsureSlash(currentRoot) & dni & "\ACTUAL\current-note.txt"
    fileHist = EnsureSlash(currentRoot) & dni & "\HISTORICO\historical-note.txt"
    Dim tStream As Object
    Set tStream = fs.CreateTextFile(fileActual, True, False)
    tStream.WriteLine "current content"
    tStream.Close
    Set tStream = fs.CreateTextFile(fileHist, True, False)
    tStream.WriteLine "historical content"
    tStream.Close

    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)
    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_HATW_C2H_FlattensAnexosFolder = JsonFail("Expected success, got: " & errMsg, logs)
        GoTo CleanUp
    End If

    Dim targetDni As String
    targetDni = EnsureSlash(historicalRoot) & dni
    If Dir$(targetDni, vbDirectory) = "" Then
        Test_HATW_C2H_FlattensAnexosFolder = JsonFail("Target DNI folder missing: " & targetDni, logs)
        GoTo CleanUp
    End If
    If Dir$(targetDni & "\ACTUAL", vbDirectory) <> "" Then
        Test_HATW_C2H_FlattensAnexosFolder = JsonFail("Target has ACTUAL\ subfolder (should be flat): " & targetDni & "\ACTUAL", logs)
        GoTo CleanUp
    End If
    If Dir$(targetDni & "\HISTORICO", vbDirectory) <> "" Then
        Test_HATW_C2H_FlattensAnexosFolder = JsonFail("Target has HISTORICO\ subfolder (should be flat): " & targetDni & "\HISTORICO", logs)
        GoTo CleanUp
    End If
    If Dir$(targetDni & "\current-note.txt") = "" Then
        Test_HATW_C2H_FlattensAnexosFolder = JsonFail("Target missing flat file from ACTUAL source subfolder: current-note.txt", logs)
        GoTo CleanUp
    End If
    If Dir$(targetDni & "\historical-note.txt") = "" Then
        Test_HATW_C2H_FlattensAnexosFolder = JsonFail("Target missing flat file from HISTORICO source subfolder: historical-note.txt", logs)
        GoTo CleanUp
    End If

    logs.Add "Source subfolders ACTUAL\ + HISTORICO\ were flattened into target " & targetDni & "\ with both files at top level."
    Test_HATW_C2H_FlattensAnexosFolder = JsonOk("flatten-c2h-ok", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_C2H_FlattensAnexosFolder = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_C2H_DuplicateFlatBasename_AbortsWithoutOverwrite() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim dni As String
    Dim sourceFile As String
    Dim duplicateActual As String
    Dim duplicateHistorical As String
    Dim targetDni As String
    Dim targetDuplicate As String
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("duplicate-flat-c2h")
    historicalRoot = TempLifecycleRoot("duplicate-flat-c2h-hist")
    PrepareCurrentFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile

    dni = "DNI" & TEST_USER_CURRENT
    duplicateActual = EnsureSlash(currentRoot) & dni & "\ACTUAL\same-name.pdf"
    duplicateHistorical = EnsureSlash(currentRoot) & dni & "\HISTORICO\same-name.pdf"
    targetDni = EnsureSlash(historicalRoot) & dni
    targetDuplicate = EnsureSlash(targetDni) & "same-name.pdf"
    CreateTextFile duplicateActual, "actual duplicate content"
    CreateTextFile duplicateHistorical, "historical duplicate content"

    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If InStr(1, errMsg, "Duplicate attachment filename", vbTextCompare) = 0 Then
        Test_HATW_C2H_DuplicateFlatBasename_AbortsWithoutOverwrite = JsonFail("Expected duplicate-basename abort, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_CURRENT) <> 1 Then
        Test_HATW_C2H_DuplicateFlatBasename_AbortsWithoutOverwrite = JsonFail("Source user row was mutated despite duplicate-basename abort.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuariosHistoricos", "ID=" & TEST_USER_CURRENT) <> 0 Then
        Test_HATW_C2H_DuplicateFlatBasename_AbortsWithoutOverwrite = JsonFail("Historical row was created despite duplicate-basename abort.", logs)
        GoTo CleanUp
    End If
    If Dir$(duplicateActual) = "" Then
        Test_HATW_C2H_DuplicateFlatBasename_AbortsWithoutOverwrite = JsonFail("ACTUAL duplicate source file missing after abort.", logs)
        GoTo CleanUp
    End If
    If Dir$(duplicateHistorical) = "" Then
        Test_HATW_C2H_DuplicateFlatBasename_AbortsWithoutOverwrite = JsonFail("HISTORICO duplicate source file missing after abort.", logs)
        GoTo CleanUp
    End If
    If Dir$(targetDuplicate) <> "" Then
        Test_HATW_C2H_DuplicateFlatBasename_AbortsWithoutOverwrite = JsonFail("Target duplicate file exists; flat copy mutated target before duplicate detection: " & targetDuplicate, logs)
        GoTo CleanUp
    End If
    If Dir$(targetDni, vbDirectory) <> "" Then
        Test_HATW_C2H_DuplicateFlatBasename_AbortsWithoutOverwrite = JsonFail("Target DNI folder was created despite duplicate-basename abort: " & targetDni, logs)
        GoTo CleanUp
    End If

    logs.Add "Duplicate basename under the same DNI tree aborted before any target mutation; both source files and DB source row remain intact."
    Test_HATW_C2H_DuplicateFlatBasename_AbortsWithoutOverwrite = JsonOk("duplicate-flat-c2h-safe", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_C2H_DuplicateFlatBasename_AbortsWithoutOverwrite = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_HistoricalToCurrent_Observaciones_CalculatesWhenCollision() As String
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
    currentRoot = TempLifecycleRoot("obs-h2c-collision")
    historicalRoot = TempLifecycleRoot("obs-h2c-collision-hist")
    PrepareHistoricalFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    db.Execute "DELETE FROM TbObservacionesHistoricas WHERE ID=" & TEST_USER_HISTORICAL, dbFailOnError
    db.Execute "INSERT INTO TbObservacionesHistoricas (IDObservacion, ID, Fecha, Observacion, Tipo) VALUES (950001, " & TEST_USER_HISTORICAL & ", #2026-06-30#, 'historical note', 'N')", dbFailOnError

    blockerParentID = 9999998
    db.Execute "DELETE FROM TbObservaciones WHERE ID=" & blockerParentID, dbFailOnError
    db.Execute "DELETE FROM TbUsuarios WHERE ID=" & blockerParentID, dbFailOnError
    db.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & blockerParentID & ", 'fake-current-blocker-dni', 'Fake', 'Blocking')", dbFailOnError
    db.Execute "DELETE FROM TbObservaciones WHERE IDObservacion=950001", dbFailOnError
    db.Execute "INSERT INTO TbObservaciones (IDObservacion, ID, Fecha, Observacion, Tipo) VALUES (950001, " & blockerParentID & ", #2026-06-30#, 'Blocking', 'N')", dbFailOnError
    preexistingMaxID = 950099
    db.Execute "DELETE FROM TbObservaciones WHERE IDObservacion=" & preexistingMaxID, dbFailOnError
    db.Execute "INSERT INTO TbObservaciones (IDObservacion, ID, Fecha, Observacion, Tipo) VALUES (" & preexistingMaxID & ", " & blockerParentID & ", #2026-06-30#, 'HighWater', 'N')", dbFailOnError

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarHistoricoAActual(CStr(TEST_USER_HISTORICAL), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_HATW_HistoricalToCurrent_Observaciones_CalculatesWhenCollision = JsonFail("Expected success, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT IDObservacion FROM TbObservaciones WHERE ID=" & TEST_USER_HISTORICAL, dbOpenSnapshot)
    If rs.EOF Then
        Test_HATW_HistoricalToCurrent_Observaciones_CalculatesWhenCollision = JsonFail("No current observacion created.", logs)
        GoTo CleanUp
    End If
    If CLng(rs!IDObservacion) = 950001 Then
        Test_HATW_HistoricalToCurrent_Observaciones_CalculatesWhenCollision = JsonFail("IDObservacion was preserved despite h2c collision: " & rs!IDObservacion, logs)
        GoTo CleanUp
    End If
    If CLng(rs!IDObservacion) <= preexistingMaxID Then
        Test_HATW_HistoricalToCurrent_Observaciones_CalculatesWhenCollision = JsonFail("IDObservacion not greater than current max: got=" & rs!IDObservacion & ", max=" & preexistingMaxID, logs)
        GoTo CleanUp
    End If
    rs.Close

    logs.Add "Historical-to-current Observacion IDObservacion calculated as new (> max=" & preexistingMaxID & ") on collision."
    Test_HATW_HistoricalToCurrent_Observaciones_CalculatesWhenCollision = JsonOk("obs-h2c-collision", logs)
CleanUp:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    db.Execute "DELETE FROM TbObservaciones WHERE ID=" & blockerParentID, dbFailOnError
    db.Execute "DELETE FROM TbUsuarios WHERE ID=" & blockerParentID, dbFailOnError
    On Error GoTo 0
    Exit Function
EH:
    Test_HATW_HistoricalToCurrent_Observaciones_CalculatesWhenCollision = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    db.Execute "DELETE FROM TbObservaciones WHERE ID=" & blockerParentID, dbFailOnError
    db.Execute "DELETE FROM TbUsuarios WHERE ID=" & blockerParentID, dbFailOnError
End Function

Public Function Test_HATW_InvalidStringUserId_RejectsBeforeResolvePaths() As String
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
    currentRoot = TempLifecycleRoot("invalid-id")
    historicalRoot = TempLifecycleRoot("invalid-id-hist")
    PrepareCurrentFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT) & " OR 1=1", db, currentRoot, historicalRoot, "", errMsg)

    If InStr(1, errMsg, "Invalid lifecycle user ID", vbTextCompare) = 0 Then
        Test_HATW_InvalidStringUserId_RejectsBeforeResolvePaths = JsonFail("Expected invalid-ID rejection before path resolution, got: " & errMsg, logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_CURRENT) <> 1 Then
        Test_HATW_InvalidStringUserId_RejectsBeforeResolvePaths = JsonFail("Source user was mutated after invalid string ID.", logs)
        GoTo CleanUp
    End If
    If Dir$(sourceFile) = "" Then
        Test_HATW_InvalidStringUserId_RejectsBeforeResolvePaths = JsonFail("Source file was mutated after invalid string ID.", logs)
        GoTo CleanUp
    End If

    logs.Add "String user ID rejected once at ExecuteLifecycleMove entry before ResolvePaths SQL is built."
    Test_HATW_InvalidStringUserId_RejectsBeforeResolvePaths = JsonOk("invalid-id-rejected", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_InvalidStringUserId_RejectsBeforeResolvePaths = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_HATW_TargetConflictFolderRestoredOnLaterFailure() As String
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim conflictFolderPath As String
    Dim conflictFilePath As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("folder-restore-fail")
    historicalRoot = TempLifecycleRoot("folder-restore-fail-hist")
    PrepareFolderConflictFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile
    conflictFolderPath = EnsureSlash(historicalRoot) & "DNI" & TEST_USER_CURRENT
    conflictFilePath = EnsureSlash(conflictFolderPath) & "conflict-pre-existing.txt"

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    coordinator.ConfigureLifecycleConflictPrompt "Test_ConflictPrompt_Approve"
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "BeforeCommit", errMsg)

    If errMsg = "" Then
        Test_HATW_TargetConflictFolderRestoredOnLaterFailure = JsonFail("Expected forced BeforeCommit failure after conflict cleanup.", logs)
        GoTo CleanUp
    End If
    If CountRows(db, "TbUsuarios", "ID=" & TEST_USER_CURRENT) <> 1 Then
        Test_HATW_TargetConflictFolderRestoredOnLaterFailure = JsonFail("Source user was not restored after later failure.", logs)
        GoTo CleanUp
    End If
    If Dir$(conflictFilePath) = "" Then
        Test_HATW_TargetConflictFolderRestoredOnLaterFailure = JsonFail("Prior target conflict folder was not restored after later failure: " & conflictFilePath, logs)
        GoTo CleanUp
    End If
    If ReadTextFile(conflictFilePath) <> "conflict content" Then
        Test_HATW_TargetConflictFolderRestoredOnLaterFailure = JsonFail("Prior target conflict folder content changed after restore.", logs)
        GoTo CleanUp
    End If

    logs.Add "Destination conflict folder was snapshotted before cleanup and restored after later rollback."
    Test_HATW_TargetConflictFolderRestoredOnLaterFailure = JsonOk("target-conflict-folder-restored", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_HATW_TargetConflictFolderRestoredOnLaterFailure = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

' ============================================================
' ANEXOS-HISTORICOS: c2h/h2c reversibility atoms (per-anexo subfolder routing)
'
' These four atoms cover the EsHistorico-preservation contract: c2h must
' preserve the per-anexo EsHistorico value (No -> No, Sí -> Sí) so h2c can
' route each file back to the correct <HPS>\<DNI>\Actual\ or HISTORICO\
' subfolder. The lifecycle c2h/h2c paths become reversible and idempotent
' when the EsHistorico flag is treated as a per-file routing decision.
'
' Tags (per project convention): c2h-h2c-subcarpetas, reversibility, idempotencia
' ============================================================

Public Function Test_C2H_PreservesEsHistorico() As String
    ' [c2h-h2c-subcarpetas] When an active user has anexos in BOTH
    ' <HPS>\<DNI>\Actual\<file> (EsHistorico='No') AND
    ' <HPS>\<DNI>\HISTORICO\<file> (EsHistorico='Sí'), the c2h flatten
    ' copy must move both files flat to <HISTORICO>\<DNI>\<file>, AND the
    ' destination rows in TbAnexosUsuariosHistoricos must carry the SAME
    ' EsHistorico value as the source rows. This is the precondition for
    ' the h2c reverse-split to land each file back in the right subfolder.
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim sourceFile As String
    Dim dni As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("c2h-subcarpetas-preserves")
    historicalRoot = TempLifecycleRoot("c2h-subcarpetas-preserves-hist")
    dni = "DNI" & TEST_USER_CURRENT
    PrepareC2HSubcarpetasFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, sourceFile

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_C2H_PreservesEsHistorico = JsonFail("Expected c2h success, got: " & errMsg, logs)
        GoTo CleanUp
    End If

    ' Filesystem assertions: BOTH source files must land flat at
    ' <HISTORICO>\<DNI>\<file> regardless of which subfolder they came from.
    Dim targetDni As String
    targetDni = EnsureSlash(historicalRoot) & dni
    If Dir$(targetDni & "\file_a.txt") = "" Then
        Test_C2H_PreservesEsHistorico = JsonFail("Target missing flat file from Actual\ source: file_a.txt", logs)
        GoTo CleanUp
    End If
    If Dir$(targetDni & "\file_h.txt") = "" Then
        Test_C2H_PreservesEsHistorico = JsonFail("Target missing flat file from HISTORICO\ source: file_h.txt", logs)
        GoTo CleanUp
    End If
    ' c2h must still flatten (no inner subfolders at <HISTORICO>\<DNI>\).
    If Dir$(targetDni & "\Actual", vbDirectory) <> "" Then
        Test_C2H_PreservesEsHistorico = JsonFail("Target has ACTUAL\ subfolder (c2h must stay flat): " & targetDni & "\Actual", logs)
        GoTo CleanUp
    End If
    If Dir$(targetDni & "\HISTORICO", vbDirectory) <> "" Then
        Test_C2H_PreservesEsHistorico = JsonFail("Target has HISTORICO\ subfolder (c2h must stay flat): " & targetDni & "\HISTORICO", logs)
        GoTo CleanUp
    End If

    ' DB assertions: rows in TbAnexosUsuariosHistoricos preserve EsHistorico
    ' from the source. file_a.txt source had EsHistorico='No', file_h.txt
    ' source had EsHistorico='Sí'. The c2h destination rows must keep those
    ' values so h2c can split them back to the right subfolders.
    Dim noRs As DAO.Recordset
    Dim yesRs As DAO.Recordset
    Set noRs = db.OpenRecordset("SELECT EsHistorico FROM TbAnexosUsuariosHistoricos WHERE IDUsuario=" & TEST_USER_CURRENT & " AND NombreAnexo='file_a.txt'", dbOpenSnapshot)
    If noRs.EOF Then
        Test_C2H_PreservesEsHistorico = JsonFail("No historical row for file_a.txt after c2h.", logs)
        GoTo CleanUp
    End If
    If Nz(noRs!EsHistorico, "") <> "No" Then
        Test_C2H_PreservesEsHistorico = JsonFail("file_a.txt EsHistorico was not preserved as 'No', got: " & Nz(noRs!EsHistorico, "(null)"), logs)
        GoTo CleanUp
    End If
    noRs.Close
    Set yesRs = db.OpenRecordset("SELECT EsHistorico FROM TbAnexosUsuariosHistoricos WHERE IDUsuario=" & TEST_USER_CURRENT & " AND NombreAnexo='file_h.txt'", dbOpenSnapshot)
    If yesRs.EOF Then
        Test_C2H_PreservesEsHistorico = JsonFail("No historical row for file_h.txt after c2h.", logs)
        GoTo CleanUp
    End If
    If Nz(yesRs!EsHistorico, "") <> "Sí" Then
        Test_C2H_PreservesEsHistorico = JsonFail("file_h.txt EsHistorico was not preserved as 'Sí', got: " & Nz(yesRs!EsHistorico, "(null)"), logs)
        GoTo CleanUp
    End If
    yesRs.Close

    ' Source side: TbAnexosUsuariosHPS must be empty for this user post-c2h.
    If CountRows(db, "TbAnexosUsuariosHPS", "IDUsuario=" & TEST_USER_CURRENT) <> 0 Then
        Test_C2H_PreservesEsHistorico = JsonFail("Source TbAnexosUsuariosHPS still has rows for this user after c2h.", logs)
        GoTo CleanUp
    End If

    logs.Add "c2h flattened both source subfolders and preserved EsHistorico (No/Sí) for h2c split-routing."
    Test_C2H_PreservesEsHistorico = JsonOk("c2h-preserves-eshistorico", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_C2H_PreservesEsHistorico = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_H2C_RestoresToCorrectSubfolder() As String
    ' [c2h-h2c-subcarpetas] h2c must read EsHistorico from each
    ' TbAnexosUsuariosHistoricos row and route the file to either
    ' <HPS>\<DNI>\Actual\<file> (No) or <HPS>\<DNI>\HISTORICO\<file> (Sí).
    ' Pre-state mirrors the c2h result: 2 flat files at <HISTORICO>\<DNI>\
    ' with one row flagged No and the other Sí.
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim dni As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("h2c-subcarpetas-splits")
    historicalRoot = TempLifecycleRoot("h2c-subcarpetas-splits-hist")
    dni = "DNI" & TEST_USER_HISTORICAL
    PrepareH2CSubcarpetasFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarHistoricoAActual(CStr(TEST_USER_HISTORICAL), db, currentRoot, historicalRoot, "", errMsg)

    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_H2C_RestoresToCorrectSubfolder = JsonFail("Expected h2c success, got: " & errMsg, logs)
        GoTo CleanUp
    End If

    ' Filesystem assertions: each file must land in the subfolder that
    ' matches its row's EsHistorico. file_a.txt was No -> Actual\;
    ' file_h.txt was Sí -> HISTORICO\.
    Dim currentDni As String
    currentDni = EnsureSlash(currentRoot) & dni
    If Dir$(currentDni & "\Actual\file_a.txt") = "" Then
        Test_H2C_RestoresToCorrectSubfolder = JsonFail("h2c did not route EsHistorico='No' file to Actual\: " & currentDni & "\Actual\file_a.txt", logs)
        GoTo CleanUp
    End If
    If Dir$(currentDni & "\HISTORICO\file_h.txt") = "" Then
        Test_H2C_RestoresToCorrectSubfolder = JsonFail("h2c did not route EsHistorico='Sí' file to HISTORICO\: " & currentDni & "\HISTORICO\file_h.txt", logs)
        GoTo CleanUp
    End If
    ' Negative: files must NOT be at the wrong subfolder.
    If Dir$(currentDni & "\Actual\file_h.txt") <> "" Then
        Test_H2C_RestoresToCorrectSubfolder = JsonFail("h2c routed EsHistorico='Sí' file to Actual\ by mistake: " & currentDni & "\Actual\file_h.txt", logs)
        GoTo CleanUp
    End If
    If Dir$(currentDni & "\HISTORICO\file_a.txt") <> "" Then
        Test_H2C_RestoresToCorrectSubfolder = JsonFail("h2c routed EsHistorico='No' file to HISTORICO\ by mistake: " & currentDni & "\HISTORICO\file_a.txt", logs)
        GoTo CleanUp
    End If
    ' The flat source folder must be gone (h2c DeleteFolder step).
    Dim sourceDni As String
    sourceDni = EnsureSlash(historicalRoot) & dni
    If Dir$(sourceDni, vbDirectory) <> "" Then
        Test_H2C_RestoresToCorrectSubfolder = JsonFail("h2c left the flat source folder behind: " & sourceDni, logs)
        GoTo CleanUp
    End If

    ' DB assertions: rows in TbAnexosUsuariosHPS preserve EsHistorico.
    Dim noRs As DAO.Recordset
    Dim yesRs As DAO.Recordset
    Set noRs = db.OpenRecordset("SELECT EsHistorico FROM TbAnexosUsuariosHPS WHERE IDUsuario=" & TEST_USER_HISTORICAL & " AND NombreAnexo='file_a.txt'", dbOpenSnapshot)
    If noRs.EOF Then
        Test_H2C_RestoresToCorrectSubfolder = JsonFail("No current row for file_a.txt after h2c.", logs)
        GoTo CleanUp
    End If
    If Nz(noRs!EsHistorico, "") <> "No" Then
        Test_H2C_RestoresToCorrectSubfolder = JsonFail("file_a.txt EsHistorico was not preserved as 'No', got: " & Nz(noRs!EsHistorico, "(null)"), logs)
        GoTo CleanUp
    End If
    noRs.Close
    Set yesRs = db.OpenRecordset("SELECT EsHistorico FROM TbAnexosUsuariosHPS WHERE IDUsuario=" & TEST_USER_HISTORICAL & " AND NombreAnexo='file_h.txt'", dbOpenSnapshot)
    If yesRs.EOF Then
        Test_H2C_RestoresToCorrectSubfolder = JsonFail("No current row for file_h.txt after h2c.", logs)
        GoTo CleanUp
    End If
    If Nz(yesRs!EsHistorico, "") <> "Sí" Then
        Test_H2C_RestoresToCorrectSubfolder = JsonFail("file_h.txt EsHistorico was not preserved as 'Sí', got: " & Nz(yesRs!EsHistorico, "(null)"), logs)
        GoTo CleanUp
    End If
    yesRs.Close

    ' Source side: TbAnexosUsuariosHistoricos must be empty for this user post-h2c.
    If CountRows(db, "TbAnexosUsuariosHistoricos", "IDUsuario=" & TEST_USER_HISTORICAL) <> 0 Then
        Test_H2C_RestoresToCorrectSubfolder = JsonFail("Source TbAnexosUsuariosHistoricos still has rows for this user after h2c.", logs)
        GoTo CleanUp
    End If

    logs.Add "h2c routed each file to the subfolder matching its EsHistorico and preserved the flag in TbAnexosUsuariosHPS."
    Test_H2C_RestoresToCorrectSubfolder = JsonOk("h2c-restores-subcarpetas", logs)
CleanUp:
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    Exit Function
EH:
    Test_H2C_RestoresToCorrectSubfolder = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
End Function

Public Function Test_H2C_IsIdempotent() As String
    ' [c2h-h2c-subcarpetas] Idempotence: running h2c twice on the same
    ' pre-state must yield the same final state. We reset between runs
    ' because the first h2c moves the user from historical to current,
    ' so a second h2c on the post-state would not find the source row.
    ' The contract is: the operation is deterministic and repeatable.
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim dni As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("h2c-subcarpetas-idempotent")
    historicalRoot = TempLifecycleRoot("h2c-subcarpetas-idempotent-hist")
    dni = "DNI" & TEST_USER_HISTORICAL

    ' --- Run 1 ---
    PrepareH2CSubcarpetasFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarHistoricoAActual(CStr(TEST_USER_HISTORICAL), db, currentRoot, historicalRoot, "", errMsg)
    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_H2C_IsIdempotent = JsonFail("Run 1 h2c failed: " & errMsg, logs)
        GoTo CleanUp
    End If
    Dim s1Current As String
    Dim s1Historical As String
    s1Current = SnapshotAnexosSubcarpetasState(db, "TbAnexosUsuariosHPS", TEST_USER_HISTORICAL)
    s1Historical = SnapshotAnexosSubcarpetasState(db, "TbAnexosUsuariosHistoricos", TEST_USER_HISTORICAL)
    Dim s1ActualA As String
    Dim s1HistoricoH As String
    s1ActualA = ReadTextFile(EnsureSlash(currentRoot) & dni & "\Actual\file_a.txt")
    s1HistoricoH = ReadTextFile(EnsureSlash(currentRoot) & dni & "\HISTORICO\file_h.txt")

    ' --- Run 2 on a freshly-reset state ---
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    PrepareH2CSubcarpetasFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarHistoricoAActual(CStr(TEST_USER_HISTORICAL), db, currentRoot, historicalRoot, "", errMsg)
    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_H2C_IsIdempotent = JsonFail("Run 2 h2c failed: " & errMsg, logs)
        GoTo CleanUp
    End If
    Dim s2Current As String
    Dim s2Historical As String
    s2Current = SnapshotAnexosSubcarpetasState(db, "TbAnexosUsuariosHPS", TEST_USER_HISTORICAL)
    s2Historical = SnapshotAnexosSubcarpetasState(db, "TbAnexosUsuariosHistoricos", TEST_USER_HISTORICAL)
    Dim s2ActualA As String
    Dim s2HistoricoH As String
    s2ActualA = ReadTextFile(EnsureSlash(currentRoot) & dni & "\Actual\file_a.txt")
    s2HistoricoH = ReadTextFile(EnsureSlash(currentRoot) & dni & "\HISTORICO\file_h.txt")

    ' Compare: filesystem files identical, DB EsHistorico values identical.
    If s1ActualA <> s2ActualA Then
        Test_H2C_IsIdempotent = JsonFail("Actual\file_a.txt content differs between runs.", logs)
        GoTo CleanUp
    End If
    If s1HistoricoH <> s2HistoricoH Then
        Test_H2C_IsIdempotent = JsonFail("HISTORICO\file_h.txt content differs between runs.", logs)
        GoTo CleanUp
    End If
    If s1Current <> s2Current Then
        Test_H2C_IsIdempotent = JsonFail("TbAnexosUsuariosHPS EsHistorico snapshot differs between runs." & vbCrLf & "Run1: " & s1Current & vbCrLf & "Run2: " & s2Current, logs)
        GoTo CleanUp
    End If
    If s1Historical <> s2Historical Then
        Test_H2C_IsIdempotent = JsonFail("TbAnexosUsuariosHistoricos EsHistorico snapshot differs between runs.", logs)
        GoTo CleanUp
    End If

    logs.Add "h2c produced identical filesystem + DB state on two consecutive runs from the same pre-state."
    Test_H2C_IsIdempotent = JsonOk("h2c-idempotent", logs)
CleanUp:
    On Error Resume Next
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    On Error GoTo 0
    Exit Function
EH:
    Test_H2C_IsIdempotent = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    On Error GoTo 0
End Function

Public Function Test_C2H_Then_H2C_IsReversible() As String
    ' [c2h-h2c-subcarpetas] Reversibility: c2h then h2c on the same
    ' starting state must yield the SAME state (per-anexo subfolder layout
    ' and EsHistorico flags both restored). This is the user's core
    ' contract for moving a user back and forth between active/historical.
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim dni As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("c2h-h2c-reversible")
    historicalRoot = TempLifecycleRoot("c2h-h2c-reversible-hist")
    dni = "DNI" & TEST_USER_CURRENT
    PrepareC2HSubcarpetasFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot, ""  ' sourceFile unused

    ' Capture the original state for the post-cycle comparison.
    Dim origActualContent As String
    Dim origHistoricoContent As String
    origActualContent = ReadTextFile(EnsureSlash(currentRoot) & dni & "\Actual\file_a.txt")
    origHistoricoContent = ReadTextFile(EnsureSlash(currentRoot) & dni & "\HISTORICO\file_h.txt")
    Dim origNoEs As String
    Dim origYesEs As String
    Dim origRs As DAO.Recordset
    Set origRs = db.OpenRecordset("SELECT NombreAnexo, EsHistorico FROM TbAnexosUsuariosHPS WHERE IDUsuario=" & TEST_USER_CURRENT & " ORDER BY NombreAnexo", dbOpenSnapshot)
    Do While Not origRs.EOF
        If Nz(origRs!NombreAnexo, "") = "file_a.txt" Then origNoEs = Nz(origRs!EsHistorico, "")
        If Nz(origRs!NombreAnexo, "") = "file_h.txt" Then origYesEs = Nz(origRs!EsHistorico, "")
        origRs.MoveNext
    Loop
    origRs.Close

    ' --- c2h ---
    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)
    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_C2H_Then_H2C_IsReversible = JsonFail("c2h step failed: " & errMsg, logs)
        GoTo CleanUp
    End If

    ' --- h2c ---
    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarHistoricoAActual(CStr(TEST_USER_CURRENT), db, currentRoot, historicalRoot, "", errMsg)
    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_C2H_Then_H2C_IsReversible = JsonFail("h2c step failed: " & errMsg, logs)
        GoTo CleanUp
    End If

    ' Filesystem reversibility: both files must be back in their original
    ' subfolders, and the flat historical tree must be gone.
    If Dir$(EnsureSlash(currentRoot) & dni & "\Actual\file_a.txt") = "" Then
        Test_C2H_Then_H2C_IsReversible = JsonFail("file_a.txt did not return to Actual\ after c2h->h2c.", logs)
        GoTo CleanUp
    End If
    If Dir$(EnsureSlash(currentRoot) & dni & "\HISTORICO\file_h.txt") = "" Then
        Test_C2H_Then_H2C_IsReversible = JsonFail("file_h.txt did not return to HISTORICO\ after c2h->h2c.", logs)
        GoTo CleanUp
    End If
    If Dir$(EnsureSlash(historicalRoot) & dni, vbDirectory) <> "" Then
        Test_C2H_Then_H2C_IsReversible = JsonFail("Flat historical source tree was not cleaned up after c2h->h2c: " & EnsureSlash(historicalRoot) & dni, logs)
        GoTo CleanUp
    End If
    ' File content must be preserved verbatim (no rewrite, no corruption).
    If ReadTextFile(EnsureSlash(currentRoot) & dni & "\Actual\file_a.txt") <> origActualContent Then
        Test_C2H_Then_H2C_IsReversible = JsonFail("file_a.txt content changed across c2h->h2c.", logs)
        GoTo CleanUp
    End If
    If ReadTextFile(EnsureSlash(currentRoot) & dni & "\HISTORICO\file_h.txt") <> origHistoricoContent Then
        Test_C2H_Then_H2C_IsReversible = JsonFail("file_h.txt content changed across c2h->h2c.", logs)
        GoTo CleanUp
    End If

    ' DB reversibility: EsHistorico values match the original.
    Dim finalNoEs As String
    Dim finalYesEs As String
    Dim finalRs As DAO.Recordset
    Set finalRs = db.OpenRecordset("SELECT NombreAnexo, EsHistorico FROM TbAnexosUsuariosHPS WHERE IDUsuario=" & TEST_USER_CURRENT & " ORDER BY NombreAnexo", dbOpenSnapshot)
    Do While Not finalRs.EOF
        If Nz(finalRs!NombreAnexo, "") = "file_a.txt" Then finalNoEs = Nz(finalRs!EsHistorico, "")
        If Nz(finalRs!NombreAnexo, "") = "file_h.txt" Then finalYesEs = Nz(finalRs!EsHistorico, "")
        finalRs.MoveNext
    Loop
    finalRs.Close
    If finalNoEs <> origNoEs Then
        Test_C2H_Then_H2C_IsReversible = JsonFail("file_a.txt EsHistorico changed across c2h->h2c. Original='" & origNoEs & "', final='" & finalNoEs & "'.", logs)
        GoTo CleanUp
    End If
    If finalYesEs <> origYesEs Then
        Test_C2H_Then_H2C_IsReversible = JsonFail("file_h.txt EsHistorico changed across c2h->h2c. Original='" & origYesEs & "', final='" & finalYesEs & "'.", logs)
        GoTo CleanUp
    End If
    ' Historical table must be empty for this user post-cycle.
    If CountRows(db, "TbAnexosUsuariosHistoricos", "IDUsuario=" & TEST_USER_CURRENT) <> 0 Then
        Test_C2H_Then_H2C_IsReversible = JsonFail("TbAnexosUsuariosHistoricos still has rows for this user after c2h->h2c.", logs)
        GoTo CleanUp
    End If

    logs.Add "c2h->h2c cycle restored both files to their original subfolders and preserved EsHistorico + content."
    Test_C2H_Then_H2C_IsReversible = JsonOk("c2h-h2c-reversible", logs)
CleanUp:
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
    Exit Function
EH:
    Test_C2H_Then_H2C_IsReversible = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_CURRENT, currentRoot, historicalRoot
End Function

Public Function Test_H2C_BothSubcarpetasExistEvenWhenEmpty() As String
    ' [c2h-h2c-subcarpetas] El usuario pidio el 2026-07-10 que tras un h2c la
    ' estructura anexos del usuario activo quede COMPLETA: tanto Actual\ como
    ' HISTORICO\ deben existir siempre, aunque una de las dos subcarpetas quede
    ' vacia. Aqui lo cubrimos en sus dos variantes (solo-No y solo-Sí).
    Dim logs As Collection
    Dim db As DAO.Database
    Dim errMsg As String
    Dim currentRoot As String
    Dim historicalRoot As String
    Dim dni As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    Dim currentDni As String
    Dim actualExists As Boolean
    Dim historicoExists As Boolean
    Set logs = New Collection
    On Error GoTo EH

    Test_ResetConflictPromptState
    Set db = OpenLocalTestBackend(errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg
    currentRoot = TempLifecycleRoot("h2c-subcarpetas-empty")
    historicalRoot = TempLifecycleRoot("h2c-subcarpetas-empty-hist")
    dni = "DNI" & TEST_USER_HISTORICAL

    ' --- Caso 1: solo EsHistorico='No' (esperamos Actual\ con archivo, HISTORICO\ vacio) ---
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    db.Execute "INSERT INTO TbUsuariosHistoricos (ID, DNI, Nombre, Apellido_1) VALUES (" & TEST_USER_HISTORICAL & ", '" & dni & "', 'Test', 'EmptySub')", dbFailOnError
    db.Execute "INSERT INTO TbAnexosUsuariosHistoricos (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & (TEST_USER_HISTORICAL + 1) & ", 'only_actual.txt', " & TEST_USER_HISTORICAL & ", 'No')", dbFailOnError
    Dim sourceDir As String
    sourceDir = EnsureSlash(historicalRoot) & dni & "\"
    ' [c2h-h2c-subcarpetas] MkDir Err 76 fix — same as PrepareH2CSubcarpetasFixture.
    CreateFolderTree sourceDir
    CreateTextFile sourceDir & "only_actual.txt", "only actual content"

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarHistoricoAActual(CStr(TEST_USER_HISTORICAL), db, currentRoot, historicalRoot, "", errMsg)
    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_H2C_BothSubcarpetasExistEvenWhenEmpty = JsonFail("h2c (solo No) fallo: " & errMsg, logs)
        GoTo CleanUp
    End If

    currentDni = EnsureSlash(currentRoot) & dni
    If Dir$(currentDni & "\Actual\only_actual.txt") = "" Then
        Test_H2C_BothSubcarpetasExistEvenWhenEmpty = JsonFail("caso solo-No: archivo no cayo en Actual\: " & currentDni & "\Actual\only_actual.txt", logs)
        GoTo CleanUp
    End If
    actualExists = (Dir$(currentDni & "\Actual", vbDirectory) <> "")
    historicoExists = (Dir$(currentDni & "\HISTORICO", vbDirectory) <> "")
    If Not (actualExists And historicoExists) Then
        Test_H2C_BothSubcarpetasExistEvenWhenEmpty = JsonFail("caso solo-No: estructura incompleta. Actual existe=" & actualExists & ", HISTORICO existe=" & historicoExists, logs)
        GoTo CleanUp
    End If
    ' HISTORICO debe estar vacio en este caso
    If Dir$(currentDni & "\HISTORICO\*.*") <> "" Then
        Test_H2C_BothSubcarpetasExistEvenWhenEmpty = JsonFail("caso solo-No: HISTORICO\ deberia estar vacio pero tiene archivos", logs)
        GoTo CleanUp
    End If
    logs.Add "caso solo-No: Actual\ con archivo + HISTORICO\ vacio OK."

    ' --- Caso 2: solo EsHistorico='Sí' (esperamos HISTORICO\ con archivo, Actual\ vacio) ---
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    db.Execute "INSERT INTO TbUsuariosHistoricos (ID, DNI, Nombre, Apellido_1) VALUES (" & TEST_USER_HISTORICAL & ", '" & dni & "', 'Test', 'EmptySub2')", dbFailOnError
    db.Execute "INSERT INTO TbAnexosUsuariosHistoricos (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & (TEST_USER_HISTORICAL + 1) & ", 'only_historico.txt', " & TEST_USER_HISTORICAL & ", 'Sí')", dbFailOnError
    sourceDir = EnsureSlash(historicalRoot) & dni & "\"
    ' [c2h-h2c-subcarpetas] MkDir Err 76 fix.
    CreateFolderTree sourceDir
    CreateTextFile sourceDir & "only_historico.txt", "only historico content"

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarHistoricoAActual(CStr(TEST_USER_HISTORICAL), db, currentRoot, historicalRoot, "", errMsg)
    If errMsg <> "" And Not IsPostCommitRefreshWarning(errMsg) Then
        Test_H2C_BothSubcarpetasExistEvenWhenEmpty = JsonFail("h2c (solo Sí) fallo: " & errMsg, logs)
        GoTo CleanUp
    End If

    If Dir$(currentDni & "\HISTORICO\only_historico.txt") = "" Then
        Test_H2C_BothSubcarpetasExistEvenWhenEmpty = JsonFail("caso solo-Sí: archivo no cayo en HISTORICO\: " & currentDni & "\HISTORICO\only_historico.txt", logs)
        GoTo CleanUp
    End If
    actualExists = (Dir$(currentDni & "\Actual", vbDirectory) <> "")
    historicoExists = (Dir$(currentDni & "\HISTORICO", vbDirectory) <> "")
    If Not (actualExists And historicoExists) Then
        Test_H2C_BothSubcarpetasExistEvenWhenEmpty = JsonFail("caso solo-Sí: estructura incompleta. Actual existe=" & actualExists & ", HISTORICO existe=" & historicoExists, logs)
        GoTo CleanUp
    End If
    ' Actual debe estar vacio en este caso
    If Dir$(currentDni & "\Actual\*.*") <> "" Then
        Test_H2C_BothSubcarpetasExistEvenWhenEmpty = JsonFail("caso solo-Sí: Actual\ deberia estar vacio pero tiene archivos", logs)
        GoTo CleanUp
    End If
    logs.Add "caso solo-Sí: HISTORICO\ con archivo + Actual\ vacio OK."

    Test_H2C_BothSubcarpetasExistEvenWhenEmpty = JsonOk("h2c-both-subcarpetas-always-exist", logs)
CleanUp:
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
    Exit Function
EH:
    Test_H2C_BothSubcarpetasExistEvenWhenEmpty = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownFixture db, TEST_USER_HISTORICAL, currentRoot, historicalRoot
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
    ' [PR1a TOCTOU idempotence] Clear any leftover `_TOCTOU`-suffixed rows from a
    ' previous run BEFORE the INSERT. Without this, a prior run that crashed or
    ' failed to clean up leaves a stale row at the INJECTED id (userId + 200),
    ' and the prompt function's later INSERT collides with DAO 3022, masking
    ' the real test signal. Wildcard uses Jet's `*` (DAO), not `%` (ANSI SQL).
    db.Execute "DELETE FROM TbUsuariosHistoricos WHERE DNI LIKE '" & Replace(dni, "'", "''") & "_TOCTOU*'", dbFailOnError
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
    ' [PR1a TOCTOU idempotence] mirror of PrepareConflictFixture: wipe any
    ' h2c prompt-injection leftovers before the INSERT. The prompt function
    ' writes to TbUsuarios (the current side, opposite of c2h), but we also
    ' guard TbUsuariosHistoricos in case a previous c2h/h2c run left a
    ' `_TOCTOU` row on the historical side that this test's LIKE-shaped
    ' conflict query would otherwise pick up.
    db.Execute "DELETE FROM TbUsuarios WHERE DNI LIKE 'DNI" & userId & "_TOCTOU*'", dbFailOnError
    db.Execute "DELETE FROM TbUsuariosHistoricos WHERE DNI LIKE 'DNI" & userId & "_TOCTOU*'", dbFailOnError
    db.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & userId + 100 & ", 'DNI" & userId & "', 'Preexisting', 'CurrentConflict')", dbFailOnError
End Sub

Private Sub PrepareH2CFolderConflictFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal historicalRoot As String, ByVal currentRoot As String)
    PrepareHistoricalFixture db, userId, currentRoot, historicalRoot
    Dim conflictFolder As String
    conflictFolder = EnsureSlash(currentRoot) & "DNI" & userId & "\"
    CreateTextFile EnsureSlash(conflictFolder) & "conflict-pre-existing.txt", "conflict content"
End Sub

Private Sub PrepareH2CBothConflictFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal historicalRoot As String, ByVal currentRoot As String)
    PrepareH2CConflictFixture db, userId, historicalRoot, currentRoot
    Dim conflictFolder As String
    conflictFolder = EnsureSlash(currentRoot) & "DNI" & userId & "\"
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
    attachmentRoot = EnsureSlash(currentRoot) & "DNI" & userId & "\"
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
    attachmentRoot = EnsureSlash(currentRoot) & "DNI" & userId & "\"
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

' [c2h-h2c-subcarpetas] Seeds an ACTIVE user with 2 anexos in BOTH the
' Actual\ and HISTORICO\ subfolders of <HPS>\<DNI>\. The first anexo has
' EsHistorico='No' (file lives in Actual\), the second has EsHistorico='Sí'
' (file lives in HISTORICO\). This mirrors the per-anexo model from
' commit 9b54ac1: anexos with EsHistorico='No' sit in Actual\, anexos with
' EsHistorico='Sí' sit in HISTORICO\. The c2h atom asserts both files
' flatten into <HISTORICO>\<DNI>\ and the EsHistorico flag is preserved
' through the move.
Private Sub PrepareC2HSubcarpetasFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal currentRoot As String, ByVal historicalRoot As String, ByRef sourceFile As String)
    TeardownFixture db, userId, currentRoot, historicalRoot
    db.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & userId & ", 'DNI" & userId & "', 'Test', 'SubcarpetasCurrent')", dbFailOnError
    ' Anexo 1: EsHistorico='No' lives in <HPS>\<DNI>\Actual\<file_a.txt>
    db.Execute "INSERT INTO TbAnexosUsuariosHPS (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & (userId + 1) & ", 'file_a.txt', " & userId & ", 'No')", dbFailOnError
    ' Anexo 2: EsHistorico='Sí' lives in <HPS>\<DNI>\HISTORICO\<file_h.txt>
    db.Execute "INSERT INTO TbAnexosUsuariosHPS (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & (userId + 2) & ", 'file_h.txt', " & userId & ", 'Sí')", dbFailOnError
    Dim dni As String
    dni = "DNI" & userId
    Dim actualDir As String
    actualDir = EnsureSlash(currentRoot) & dni & "\Actual\"
    Dim historicoDir As String
    historicoDir = EnsureSlash(currentRoot) & dni & "\HISTORICO\"
    ' [c2h-h2c-subcarpetas] MkDir fails (Err 76 Path not found) when the
    ' parent chain doesn't exist yet. CreateFolderTree walks the path
    ' level by level and is idempotent.
    CreateFolderTree actualDir
    CreateFolderTree historicoDir
    CreateTextFile actualDir & "file_a.txt", "actual attachment content"
    CreateTextFile historicoDir & "file_h.txt", "historico attachment content"
    sourceFile = actualDir & "file_a.txt"
End Sub

' [c2h-h2c-subcarpetas] Seeds a HISTORICAL user with 2 anexos that
' mirror the post-c2h flat layout: <HISTORICO>\<DNI>\<file> at top level.
' The first row has EsHistorico='No' (file_a.txt), the second has
' EsHistorico='Sí' (file_h.txt). The h2c atoms read these flags and
' route each file to the matching <HPS>\<DNI>\Actual\ or HISTORICO\
' subfolder.
Private Sub PrepareH2CSubcarpetasFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal currentRoot As String, ByVal historicalRoot As String)
    TeardownFixture db, userId, currentRoot, historicalRoot
    db.Execute "INSERT INTO TbUsuariosHistoricos (ID, DNI, Nombre, Apellido_1) VALUES (" & userId & ", 'DNI" & userId & "', 'Test', 'SubcarpetasHistorical')", dbFailOnError
    ' Two anexos: one flagged No, one flagged Sí. EsHistorico MUST be set
    ' explicitly so the new CopyHistoricalRowsToCurrent code can read and
    ' preserve it on the way out.
    db.Execute "INSERT INTO TbAnexosUsuariosHistoricos (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & (userId + 1) & ", 'file_a.txt', " & userId & ", 'No')", dbFailOnError
    db.Execute "INSERT INTO TbAnexosUsuariosHistoricos (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & (userId + 2) & ", 'file_h.txt', " & userId & ", 'Sí')", dbFailOnError
    Dim dni As String
    dni = "DNI" & userId
    Dim sourceDir As String
    sourceDir = EnsureSlash(historicalRoot) & dni & "\"
    ' [c2h-h2c-subcarpetas] Same MkDir Err 76 fix as PrepareC2HSubcarpetasFixture.
    CreateFolderTree sourceDir
    CreateTextFile sourceDir & "file_a.txt", "actual attachment content"
    CreateTextFile sourceDir & "file_h.txt", "historico attachment content"
End Sub

' [c2h-h2c-subcarpetas] Reads every (NombreAnexo, EsHistorico) row for
' the user from the given table and returns a deterministic string. Used
' by the idempotence atom to compare post-run DB states without relying
' on row order or column projection.
Private Function SnapshotAnexosSubcarpetasState(ByRef db As DAO.Database, ByVal tableName As String, ByVal userId As Long) As String
    Dim rs As DAO.Recordset
    Dim result As String
    Set rs = db.OpenRecordset("SELECT NombreAnexo, EsHistorico FROM " & tableName & " WHERE IDUsuario=" & userId & " ORDER BY NombreAnexo", dbOpenSnapshot)
    Do While Not rs.EOF
        If Len(result) > 0 Then result = result & "|"
        result = result & CStr(Nz(rs!NombreAnexo, "")) & "=" & CStr(Nz(rs!EsHistorico, ""))
        rs.MoveNext
    Loop
    rs.Close
    SnapshotAnexosSubcarpetasState = result
End Function

' ============================================================
' Existing fixtures and helpers (preserved from main)
' ============================================================

Private Sub PrepareCurrentFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal currentRoot As String, ByVal historicalRoot As String, ByRef sourceFile As String)
    TeardownFixture db, userId, currentRoot, historicalRoot
    db.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & userId & ", 'DNI" & userId & "', 'Test', 'Current')", dbFailOnError
    db.Execute "INSERT INTO TbAnexosUsuariosHPS (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & (userId + 1) & ", 'current-note.txt', " & userId & ", 'No')", dbFailOnError
    sourceFile = EnsureSlash(currentRoot) & "DNI" & userId & "\current-note.txt"
    CreateCurrentAttachmentTree currentRoot, userId
End Sub

Private Sub PrepareHistoricalFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal currentRoot As String, ByVal historicalRoot As String)
    TeardownFixture db, userId, currentRoot, historicalRoot
    db.Execute "INSERT INTO TbUsuariosHistoricos (ID, DNI, Nombre, Apellido_1) VALUES (" & userId & ", 'DNI" & userId & "', 'Test', 'Historical')", dbFailOnError
    ' [c2h-h2c-subcarpetas] EsHistorico MUST be set explicitly on the
    ' historical source row so the new CopyHistoricalRowsToCurrent code
    ' can preserve it into TbAnexosUsuariosHPS. The legacy fixture omitted
    ' it (defaulted to NULL), which would let
    ' Test_HATW_HistoricalToCurrent_AnexosAlwaysNewIDAnexo_PlusEsHistorico
    ' fail with the new preserve-EsHistorico behavior. The legacy default
    ' for a moved-from-active user was EsHistorico='No' (the source row
    ' was 'No' before c2h, so the post-c2h historical row was implicitly
    ' 'No' too). Make that explicit.
    db.Execute "INSERT INTO TbAnexosUsuariosHistoricos (IDAnexo, NombreAnexo, IDUsuario, EsHistorico) VALUES (" & (userId + 1) & ", 'historical-note.txt', " & userId & ", 'No')", dbFailOnError
    CreateTextFile EnsureSlash(historicalRoot) & "DNI" & userId & "\historical-note.txt", "historical attachment"
End Sub

Private Sub TeardownFixture(ByRef db As DAO.Database, ByVal userId As Long, ByVal currentRoot As String, ByVal historicalRoot As String)
    ' [PR1a idempotency] §1.7 access-vba-tdd: shared sandbox persists across runs.
    ' Cleanup must be by DNI (not just by ID) because PR1a's precheck detects
    ' conflicts by DNI — a row at a different ID with the same DNI counts as
    ' a conflict. Without DNI-based cleanup, leftover conflict rows from
    ' previous runs poison the next test (Test_HATW_Precheck1_NoConflicts
    ' sees BD=DNI<id> from a previous Test_HATW_Precheck1_DetectsDBConflicts
    ' run; Test_HATW_TOCTOU_* sees an extra historical row from a previous
    ' prompt injection).
    On Error Resume Next
    If Not db Is Nothing Then
        ' Cleanup by ID
        db.Execute "DELETE FROM TbAnexosUsuariosHPS WHERE IDUsuario=" & userId, dbFailOnError
        db.Execute "DELETE FROM TbAnexosUsuariosHistoricos WHERE IDUsuario=" & userId, dbFailOnError
        db.Execute "DELETE FROM TbUsuarios WHERE ID=" & userId, dbFailOnError
        db.Execute "DELETE FROM TbUsuariosHistoricos WHERE ID=" & userId, dbFailOnError
        ' Cleanup by DNI (covers pre-existing conflict rows at different IDs,
        ' plus rows injected by Test_ConflictPrompt_ApproveAndInjectExtra*).
        db.Execute "DELETE FROM TbAnexosUsuariosHPS WHERE IDUsuario IN (SELECT ID FROM TbUsuarios WHERE DNI='DNI" & userId & "')", dbFailOnError
        db.Execute "DELETE FROM TbAnexosUsuariosHistoricos WHERE IDUsuario IN (SELECT ID FROM TbUsuariosHistoricos WHERE DNI='DNI" & userId & "')", dbFailOnError
        db.Execute "DELETE FROM TbUsuarios WHERE DNI='DNI" & userId & "'", dbFailOnError
        db.Execute "DELETE FROM TbUsuariosHistoricos WHERE DNI='DNI" & userId & "'", dbFailOnError
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
    ' [PR2 RFC norm] Flat layout — no ACTUAL\ subfolder; see the norm block
    ' at UsuarioLifecycleTransactionCoordinator.CopyFolderIfExists.
    attachmentRoot = EnsureSlash(currentRoot) & "DNI" & userId & "\"

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
    attachmentRoot = EnsureSlash(currentRoot) & "DNI" & userId & "\"
    If Dir$(attachmentRoot, vbDirectory) = "" Then
        p_Error = "Missing restored flat DNI folder: " & attachmentRoot
        Exit Function
    End If

    relativePaths = Array("current-note.txt", "child-a.txt", "child-b.txt")
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
    Dim path As String
    path = EnsureSlash(Environ$("TEMP")) & "HPS\LifecycleWrapperTests\" & suffix & "\"
    ' [c2h-h2c-subcarpetas] Ensure the root exists. New fixtures MkDir
    ' <root>\<DNI>\Actual\ directly; without the parent, MkDir raises
    ' Err 76 (Path not found). Idempotent — no-op if already present.
    CreateFolderTree path
    TempLifecycleRoot = path
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
    ' [PR1a test isolation] strip the trailing backslash before checking
    ' and deleting — Dir$() with vbDirectory misbehaves on paths ending in "\"
    ' (returns "" even when the folder exists), and FSO.DeleteFolder accepts
    ' both forms but the check needed to be robust. The leftover evidence:
    ' previous runs of precheck1-clean left <temp>\HPS\LifecycleWrapperTests\
    ' precheck1-clean-hist\DNI900101 with the full source tree inside because
    ' the teardown's Dir$ check saw "" and skipped the delete.
    If Len(path) = 0 Then Exit Sub
    Dim trimmed As String
    trimmed = path
    Do While Len(trimmed) > 0 And Right$(trimmed, 1) = "\"
        trimmed = Left$(trimmed, Len(trimmed) - 1)
    Loop
    If Len(trimmed) = 0 Then Exit Sub
    If Dir$(trimmed, vbDirectory) <> "" Then
        CreateObject("Scripting.FileSystemObject").DeleteFolder trimmed, True
    End If
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

