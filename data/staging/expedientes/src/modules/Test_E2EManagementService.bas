Attribute VB_Name = "Test_E2EManagementService"
Option Compare Database
Option Explicit

Public Function Test_E2EManagement_StatusFilterPendingClause() As String
    Dim p_Error As String
    Dim clause As String
    p_Error = ""

    clause = GetE2EIntegratedStatusWhereClause("pending", p_Error)
    If p_Error <> "" Then
        Test_E2EManagement_StatusFilterPendingClause = JsonOK(False, p_Error)
        Exit Function
    End If

    Test_E2EManagement_StatusFilterPendingClause = JsonOK(InStr(1, clause, "HashUltimaExportacion", vbTextCompare) > 0, "pending status clause generated")
End Function

Public Function Test_E2EManagement_PickerLeftExcludesRight() As String
    Dim p_Error As String
    Dim leftCsv As String
    p_Error = ""

    leftCsv = BuildPickerAvailableCsv("1001,1002,1003", "1002", p_Error)
    If p_Error <> "" Then
        Test_E2EManagement_PickerLeftExcludesRight = JsonOK(False, p_Error)
        Exit Function
    End If

    Test_E2EManagement_PickerLeftExcludesRight = JsonOK(leftCsv = "1001,1003", "right selection is excluded from left")
End Function

Public Function Test_E2EManagement_RemoveRightRehydratesWhenMatches() As String
    Dim p_Error As String
    Dim leftCsv As String
    p_Error = ""

    leftCsv = RehydrateFromRightRemoval(1002, True, "1001,1003", p_Error)
    If p_Error <> "" Then
        Test_E2EManagement_RemoveRightRehydratesWhenMatches = JsonOK(False, p_Error)
        Exit Function
    End If

    Test_E2EManagement_RemoveRightRehydratesWhenMatches = JsonOK(leftCsv = "1001,1003,1002", "removed item reappears on left when filter still matches")
End Function

Public Function Test_E2EManagement_RemoveRightDoesNotRehydrateWhenNoMatch() As String
    Dim p_Error As String
    Dim leftCsv As String
    p_Error = ""

    leftCsv = RehydrateFromRightRemoval(1002, False, "1001,1003", p_Error)
    If p_Error <> "" Then
        Test_E2EManagement_RemoveRightDoesNotRehydrateWhenNoMatch = JsonOK(False, p_Error)
        Exit Function
    End If

    Test_E2EManagement_RemoveRightDoesNotRehydrateWhenNoMatch = JsonOK(leftCsv = "1001,1003", "removed item stays hidden when filter no longer matches")
End Function

Public Function Test_E2EManagement_TempSelectionSessionIsolation() As String
    Dim p_Error As String
    Dim ok As Boolean
    Dim countA As Long
    Dim countB As Long
    Dim errLocal As String
    Dim cleanupErr As String
    Dim cleanupNeeded As Boolean
    Dim logs(0 To 5) As String

    ' Default: fail. Solo se override si los asserts pasan.
    Test_E2EManagement_TempSelectionSessionIsolation = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    p_Error = ""
    errLocal = ""

    logs(0) = "Arrange: setup sandbox fixtures for session isolation"
    ok = SetupE2EBatchSchemaSandbox(errLocal)
    If Not ok Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_TempSelectionSessionIsolation = BuildJsonFail("sandbox setup failed: " & errLocal, logs)
        GoTo Cleanup
    End If
    cleanupNeeded = True

    ok = SetupE2ESelectionFixture("qa.user", "S-A", "910001,910002", errLocal)
    If Not ok Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_TempSelectionSessionIsolation = BuildJsonFail("setup S-A failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    ok = SetupE2ESelectionFixture("qa.user", "S-B", "920001", errLocal)
    If Not ok Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_TempSelectionSessionIsolation = BuildJsonFail("setup S-B failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    countA = CountSelectionTempRows("qa.user", "S-A", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_TempSelectionSessionIsolation = BuildJsonFail(errLocal, logs)
        GoTo Cleanup
    End If

    countB = CountSelectionTempRows("qa.user", "S-B", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_TempSelectionSessionIsolation = BuildJsonFail(errLocal, logs)
        GoTo Cleanup
    End If

    logs(1) = "Act: counted rows for S-A=" & CStr(countA) & " and S-B=" & CStr(countB)
    logs(3) = "Assert: rows isolated by UsuarioConectado + SessionId"
    Test_E2EManagement_TempSelectionSessionIsolation = JsonOK((countA = 2 And countB = 1), "temp rows are isolated by UsuarioConectado + SessionId")
    GoTo Cleanup

HandleError:
    p_Error = "Test_E2EManagement_TempSelectionSessionIsolation: " & Err.Description
    logs(4) = p_Error
    Test_E2EManagement_TempSelectionSessionIsolation = BuildJsonFail(p_Error, logs)

Cleanup:
    On Error Resume Next
    logs(5) = "Cleanup: delete deterministic qa.user% / S-% selection fixtures"
    If cleanupNeeded Then
        cleanupErr = ""
        If Not TeardownE2EBatchSchemaSandbox(cleanupErr) Or cleanupErr <> "" Then
            p_Error = cleanupErr
            Test_E2EManagement_TempSelectionSessionIsolation = BuildJsonFail("cleanup failed: " & cleanupErr, logs)
        End If
    End If
End Function

Public Function Test_E2EManagement_SelectionFixtureUsesExplicitDbAndCardinality() As String
    Dim p_Error As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim cleanupNeeded As Boolean
    Dim db As DAO.Database
    Dim countBefore As Long
    Dim countAfter As Long
    Dim logs(0 To 6) As String

    ' Default: fail. Solo se override si los asserts pasan.
    Test_E2EManagement_SelectionFixtureUsesExplicitDbAndCardinality = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    p_Error = ""
    errLocal = ""

    logs(0) = "Schema: TbE2EExportSeleccionTemp requires UsuarioConectado, SessionId, IDExpediente; unique index covers UsuarioConectado + SessionId + IDExpediente; no DAO relationships touch this table"
    logs(1) = "Arrange: setup sandbox schema and open explicit DAO.Database"
    If Not SetupE2EBatchSchemaSandbox(errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_SelectionFixtureUsesExplicitDbAndCardinality = BuildJsonFail("sandbox setup failed: " & errLocal, logs)
        GoTo Cleanup
    End If
    cleanupNeeded = True

    Set db = GetTestDb(errLocal)
    If errLocal <> "" Or db Is Nothing Then
        p_Error = errLocal
        Test_E2EManagement_SelectionFixtureUsesExplicitDbAndCardinality = BuildJsonFail("GetTestDb failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    If Not ClearSelectionTempRows("qa.user.cardinality", "S-CARD", errLocal, db) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_SelectionFixtureUsesExplicitDbAndCardinality = BuildJsonFail("pre-clean failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    countBefore = CountSelectionTempRows("qa.user.cardinality", "S-CARD", errLocal, db)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_SelectionFixtureUsesExplicitDbAndCardinality = BuildJsonFail("count before failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(2) = "Act: seed deterministic IDs 910101 and 910102 through fixture helper with explicit DAO.Database"
    If Not SetupE2ESelectionFixture("qa.user.cardinality", "S-CARD", "910101,910102", errLocal, db) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_SelectionFixtureUsesExplicitDbAndCardinality = BuildJsonFail("fixture seed failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    countAfter = CountSelectionTempRows("qa.user.cardinality", "S-CARD", errLocal, db)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_SelectionFixtureUsesExplicitDbAndCardinality = BuildJsonFail("count after failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(3) = "Assert: before cardinality=" & CStr(countBefore) & "; after cardinality=" & CStr(countAfter)
    If countBefore = 0 And countAfter = 2 Then
        Test_E2EManagement_SelectionFixtureUsesExplicitDbAndCardinality = BuildJsonOk("selection fixture uses explicit db and cardinality changed from 0 to 2", logs)
    Else
        Test_E2EManagement_SelectionFixtureUsesExplicitDbAndCardinality = BuildJsonFail("selection fixture cardinality mismatch", logs)
    End If
    GoTo Cleanup

HandleError:
    p_Error = "Test_E2EManagement_SelectionFixtureUsesExplicitDbAndCardinality: " & Err.Description
    logs(4) = p_Error
    Test_E2EManagement_SelectionFixtureUsesExplicitDbAndCardinality = BuildJsonFail(p_Error, logs)

Cleanup:
    On Error Resume Next
    logs(5) = "Cleanup: delete deterministic qa.user.cardinality / S-CARD selection fixtures"
    If cleanupNeeded Then
        cleanupErr = ""
        If Not TeardownE2EBatchSchemaSandbox(cleanupErr) Or cleanupErr <> "" Then
            p_Error = cleanupErr
            Test_E2EManagement_SelectionFixtureUsesExplicitDbAndCardinality = BuildJsonFail("cleanup failed: " & cleanupErr, logs)
        End If
    End If
End Function

Public Function Test_E2EManagement_DestinationIsolatedByUsuarioRed() As String
    Dim p_Error As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim cleanupNeeded As Boolean
    Dim rutaA As String
    Dim rutaB As String
    Dim readA As String
    Dim readB As String
    Dim statusA As String
    Dim statusB As String
    Dim canRunA As Boolean
    Dim canRunB As Boolean
    Dim logs(0 To 4) As String

    ' Default: fail. Solo se override si los asserts pasan.
    Test_E2EManagement_DestinationIsolatedByUsuarioRed = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    p_Error = ""
    errLocal = ""

    logs(0) = "Arrange: setup sandbox and deterministic qa.user destination fixtures"
    If Not SetupE2EBatchSchemaSandbox(errLocal) Then
        p_Error = errLocal
        Test_E2EManagement_DestinationIsolatedByUsuarioRed = BuildJsonFail("sandbox setup failed: " & errLocal, logs)
        GoTo Cleanup
    End If
    cleanupNeeded = True

    rutaA = BuildWritableTempPath("qa_user_a")
    rutaB = BuildWritableTempPath("qa_user_b")

    If Not SetE2EJsonDestination("qa.user.a", rutaA, errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationIsolatedByUsuarioRed = BuildJsonFail("cannot persist destination for user A: " & errLocal, logs)
        GoTo Cleanup
    End If

    errLocal = ""
    If Not SetE2EJsonDestination("qa.user.b", rutaB, errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationIsolatedByUsuarioRed = BuildJsonFail("cannot persist destination for user B: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(1) = "Act: read persisted destination and preflight status for each UsuarioRed"
    errLocal = ""
    readA = GetE2EJsonDestination("qa.user.a", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationIsolatedByUsuarioRed = BuildJsonFail("cannot read destination for user A: " & errLocal, logs)
        GoTo Cleanup
    End If

    errLocal = ""
    readB = GetE2EJsonDestination("qa.user.b", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationIsolatedByUsuarioRed = BuildJsonFail("cannot read destination for user B: " & errLocal, logs)
        GoTo Cleanup
    End If

    errLocal = ""
    statusA = GetE2EJsonDestinationStatus("qa.user.a", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationIsolatedByUsuarioRed = BuildJsonFail("status for user A failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    errLocal = ""
    statusB = GetE2EJsonDestinationStatus("qa.user.b", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationIsolatedByUsuarioRed = BuildJsonFail("status for user B failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    errLocal = ""
    canRunA = CanRunManualE2EJsonGeneration("qa.user.a", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationIsolatedByUsuarioRed = BuildJsonFail("preflight for user A failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    errLocal = ""
    canRunB = CanRunManualE2EJsonGeneration("qa.user.b", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationIsolatedByUsuarioRed = BuildJsonFail("preflight for user B failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(2) = "Assert: destination config is isolated by exact UsuarioRed"
    If (readA = rutaA) And (readB = rutaB) And (statusA = "ok") And (statusB = "ok") And canRunA And canRunB Then
        Test_E2EManagement_DestinationIsolatedByUsuarioRed = BuildJsonOk("destination config is isolated and valid per UsuarioRed", logs)
    Else
        Test_E2EManagement_DestinationIsolatedByUsuarioRed = BuildJsonFail("destination config is not isolated or not valid per UsuarioRed", logs)
    End If
    GoTo Cleanup

HandleError:
    p_Error = "Test_E2EManagement_DestinationIsolatedByUsuarioRed: " & Err.Description
    logs(3) = p_Error
    Test_E2EManagement_DestinationIsolatedByUsuarioRed = BuildJsonFail(p_Error, logs)

Cleanup:
    On Error Resume Next
    logs(4) = "Cleanup: delete qa.user% destination rows and test-owned temp folders"
    If cleanupNeeded Then
        cleanupErr = ""
        If Not TeardownE2EJsonDestinationConfigSandbox(cleanupErr) Or cleanupErr <> "" Then
            p_Error = cleanupErr
            Test_E2EManagement_DestinationIsolatedByUsuarioRed = BuildJsonFail("cleanup failed: " & cleanupErr, logs)
        End If
    End If
    DeleteDestinationTempFolder rutaA
    DeleteDestinationTempFolder rutaB
End Function

Public Function Test_E2EManagement_DestinationMissingOrInvalidStatus() As String
    Dim p_Error As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim cleanupNeeded As Boolean
    Dim missingStatus As String
    Dim invalidStatus As String
    Dim canRunMissing As Boolean
    Dim canRunInvalid As Boolean
    Dim invalidPath As String
    Dim logs(0 To 4) As String

    ' Default: fail. Solo se override si los asserts pasan.
    Test_E2EManagement_DestinationMissingOrInvalidStatus = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    p_Error = ""
    errLocal = ""

    logs(0) = "Arrange: setup sandbox and deterministic missing/invalid destination fixtures"
    If Not SetupE2EBatchSchemaSandbox(errLocal) Then
        p_Error = errLocal
        Test_E2EManagement_DestinationMissingOrInvalidStatus = BuildJsonFail("sandbox setup failed: " & errLocal, logs)
        GoTo Cleanup
    End If
    cleanupNeeded = True

    logs(1) = "Act: evaluate missing destination status"
    missingStatus = GetE2EJsonDestinationStatus("qa.user.missing", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationMissingOrInvalidStatus = BuildJsonFail("missing status check failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    errLocal = ""
    canRunMissing = CanRunManualE2EJsonGeneration("qa.user.missing", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationMissingOrInvalidStatus = BuildJsonFail("missing preflight check failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    invalidPath = BuildInvalidTempPath("qa_user_invalid")
    errLocal = ""
    If Not SetE2EJsonDestination("qa.user.invalid", invalidPath, errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationMissingOrInvalidStatus = BuildJsonFail("cannot persist invalid path for status test: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(2) = "Act: evaluate invalid destination status"
    errLocal = ""
    invalidStatus = GetE2EJsonDestinationStatus("qa.user.invalid", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationMissingOrInvalidStatus = BuildJsonFail("invalid status check failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    errLocal = ""
    canRunInvalid = CanRunManualE2EJsonGeneration("qa.user.invalid", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationMissingOrInvalidStatus = BuildJsonFail("invalid preflight check failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(3) = "Assert: missing/invalid statuses block manual generation"
    If (missingStatus = "missing") And (invalidStatus = "invalid") And (Not canRunMissing) And (Not canRunInvalid) Then
        Test_E2EManagement_DestinationMissingOrInvalidStatus = BuildJsonOk("destination status reports missing/invalid and blocks generation", logs)
    Else
        Test_E2EManagement_DestinationMissingOrInvalidStatus = BuildJsonFail("destination status/preflight mismatch for missing or invalid destination", logs)
    End If
    GoTo Cleanup

HandleError:
    p_Error = "Test_E2EManagement_DestinationMissingOrInvalidStatus: " & Err.Description
    logs(3) = p_Error
    Test_E2EManagement_DestinationMissingOrInvalidStatus = BuildJsonFail(p_Error, logs)

Cleanup:
    On Error Resume Next
    logs(4) = "Cleanup: delete qa.user% destination rows"
    If cleanupNeeded Then
        cleanupErr = ""
        If Not TeardownE2EJsonDestinationConfigSandbox(cleanupErr) Or cleanupErr <> "" Then
            p_Error = cleanupErr
            Test_E2EManagement_DestinationMissingOrInvalidStatus = BuildJsonFail("cleanup failed: " & cleanupErr, logs)
        End If
    End If
End Function

Public Function Test_E2EManagement_DestinationUpsertSameUsuarioRed() As String
    Dim p_Error As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim cleanupNeeded As Boolean
    Dim rutaFirst As String
    Dim rutaSecond As String
    Dim readValue As String
    Dim configRows As Long
    Dim logs(0 To 5) As String

    ' Default: fail. Solo se override si los asserts pasan.
    Test_E2EManagement_DestinationUpsertSameUsuarioRed = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    p_Error = ""
    errLocal = ""

    logs(0) = "Arrange: setup sandbox and deterministic qa.user.upsert destination fixture"
    If Not SetupE2EBatchSchemaSandbox(errLocal) Then
        p_Error = errLocal
        Test_E2EManagement_DestinationUpsertSameUsuarioRed = BuildJsonFail("sandbox setup failed: " & errLocal, logs)
        GoTo Cleanup
    End If
    cleanupNeeded = True

    rutaFirst = BuildWritableTempPath("qa_user_upsert_first")
    rutaSecond = BuildWritableTempPath("qa_user_upsert_second")

    logs(1) = "Act: persist destination twice for same UsuarioRed"
    If Not SetE2EJsonDestination("qa.user.upsert", rutaFirst, errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationUpsertSameUsuarioRed = BuildJsonFail("cannot persist first destination: " & errLocal, logs)
        GoTo Cleanup
    End If

    errLocal = ""
    If Not SetE2EJsonDestination("qa.user.upsert", rutaSecond, errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationUpsertSameUsuarioRed = BuildJsonFail("cannot persist second destination: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(2) = "Assert: final route equals second value"
    readValue = GetE2EJsonDestination("qa.user.upsert", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationUpsertSameUsuarioRed = BuildJsonFail("cannot read updated destination: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(3) = "Assert: cardinality remains exactly one for same UsuarioRed"
    configRows = CountDestinationConfigRows("qa.user.upsert", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EManagement_DestinationUpsertSameUsuarioRed = BuildJsonFail("cannot count destination rows: " & errLocal, logs)
        GoTo Cleanup
    End If

    If (readValue = rutaSecond) And (configRows = 1) Then
        Test_E2EManagement_DestinationUpsertSameUsuarioRed = BuildJsonOk("destination upsert keeps second route with cardinality one", logs)
    Else
        Test_E2EManagement_DestinationUpsertSameUsuarioRed = BuildJsonFail("destination upsert mismatch: route/cardinality contract failed", logs)
    End If
    GoTo Cleanup

HandleError:
    p_Error = "Test_E2EManagement_DestinationUpsertSameUsuarioRed: " & Err.Description
    logs(4) = p_Error
    Test_E2EManagement_DestinationUpsertSameUsuarioRed = BuildJsonFail(p_Error, logs)

Cleanup:
    On Error Resume Next
    logs(5) = "Cleanup: delete qa.user% destination rows and test-owned temp folders"
    If cleanupNeeded Then
        cleanupErr = ""
        If Not TeardownE2EJsonDestinationConfigSandbox(cleanupErr) Or cleanupErr <> "" Then
            p_Error = cleanupErr
            Test_E2EManagement_DestinationUpsertSameUsuarioRed = BuildJsonFail("cleanup failed: " & cleanupErr, logs)
        End If
    End If
    DeleteDestinationTempFolder rutaFirst
    DeleteDestinationTempFolder rutaSecond
End Function

Private Function JsonOK(ByVal p_Ok As Boolean, ByVal p_Value As String) As String
    Dim logs(0 To 0) As String
    logs(0) = p_Value
    If p_Ok Then
        JsonOK = BuildJsonOk(p_Value, logs)
    Else
        JsonOK = BuildJsonFail(p_Value, logs)
    End If
End Function

Private Function BuildWritableTempPath(ByVal p_Token As String) As String
    Dim basePath As String
    basePath = Environ$("TEMP")
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    BuildWritableTempPath = basePath & "expedientes_e2e_destination_" & p_Token
    If Not fso.FolderExists(BuildWritableTempPath) Then fso.CreateFolder BuildWritableTempPath
End Function

Private Function BuildInvalidTempPath(ByVal p_Token As String) As String
    Dim basePath As String
    basePath = Environ$("TEMP")
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    BuildInvalidTempPath = basePath & "expedientes_e2e_destination_invalid_" & p_Token & "_" & Format$(Now, "yyyymmddhhnnss")
End Function

Private Sub DeleteDestinationTempFolder(ByVal p_Path As String)
    On Error Resume Next
    ' Cleanup best-effort: borra archivos dentro + carpeta. Cero dialog
    ' (RmDir + Kill nativos, no fso.DeleteFolder que puede tirar dialogs).
    If Trim$(p_Path) <> "" And Dir(p_Path, vbDirectory) <> "" Then
        Dim m_F As String
        m_F = Dir(p_Path & "\*.*")
        Do While Len(m_F) > 0
            Kill p_Path & "\" & m_F
            m_F = Dir()
        Loop
        RmDir StripTrailingSlash(p_Path)
    End If
End Sub

' Quita el trailing backslash de un path. RmDir no lo acepta.
Private Function StripTrailingSlash(ByVal p_Path As String) As String
    If Len(p_Path) > 0 And Right$(p_Path, 1) = "\" Then
        StripTrailingSlash = Left$(p_Path, Len(p_Path) - 1)
    Else
        StripTrailingSlash = p_Path
    End If
End Function

Private Function CountDestinationConfigRows(ByVal p_UsuarioRed As String, ByRef p_Error As String) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    On Error GoTo HandleError
    p_Error = ""

    If Not EnsureE2EJsonDestinationConfigSchema(p_Error) Then Exit Function

    Set db = getdb(p_Error)
    If p_Error <> "" Then Exit Function
    If db Is Nothing Then
        p_Error = "CountDestinationConfigRows: getdb returned Nothing"
        Exit Function
    End If

    Set rs = db.OpenRecordset("SELECT COUNT(*) AS Cnt FROM TbE2EJsonDestinationUserConfig WHERE UsuarioRed='" & SqlStr(p_UsuarioRed) & "'", dbOpenSnapshot)
    If Not rs.EOF Then CountDestinationConfigRows = CLng(Nz(rs.Fields("Cnt").Value, 0))

    rs.Close
    Set rs = Nothing
    Exit Function

HandleError:
    p_Error = "CountDestinationConfigRows: " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function
