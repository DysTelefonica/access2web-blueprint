Attribute VB_Name = "Test_E2EFamilyExport"
Option Compare Database
Option Explicit

Private Const FAMILY_EXPORT_TEST_ID_BASE As Long = 994012300

Private Function FamilyRootId() As Long
    FamilyRootId = FAMILY_EXPORT_TEST_ID_BASE + 1
End Function

Private Function FamilyChildId() As Long
    FamilyChildId = FAMILY_EXPORT_TEST_ID_BASE + 2
End Function

Private Function FamilySiblingId() As Long
    FamilySiblingId = FAMILY_EXPORT_TEST_ID_BASE + 3
End Function

Private Function FamilyGrandchildId() As Long
    FamilyGrandchildId = FAMILY_EXPORT_TEST_ID_BASE + 4
End Function

Private Function CycleFirstId() As Long
    CycleFirstId = FAMILY_EXPORT_TEST_ID_BASE + 10
End Function

Private Function CycleSecondId() As Long
    CycleSecondId = FAMILY_EXPORT_TEST_ID_BASE + 11
End Function

Private Function SetupE2EFamilyExportSandbox(ByRef p_Error As String) As Boolean
    p_Error = ""
    SetupE2EFamilyExportSandbox = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function
    If Not TeardownE2EFamilyExportFixture(p_Error) Then Exit Function

    SetupE2EFamilyExportSandbox = True
End Function

Private Function SeedE2EFamilyExportFixture(ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim db As DAO.Database
    p_Error = ""
    SeedE2EFamilyExportFixture = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function
    Set db = getdb(p_Error)
    If db Is Nothing Then
        If p_Error = "" Then p_Error = "SeedE2EFamilyExportFixture: getdb returned Nothing"
        Exit Function
    End If

    db.Execute "INSERT INTO TbExpedientes (IDExpediente, IDExpedientePadre, OrdinalE2E, Nemotecnico, Titulo) VALUES (" & FamilyRootId() & ", Null, 9301, 'E2E-FAMILY-ROOT', 'E2E family root');", dbFailOnError
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, IDExpedientePadre, OrdinalE2E, Nemotecnico, Titulo) VALUES (" & FamilyChildId() & ", " & FamilyRootId() & ", 9302, 'E2E-FAMILY-CHILD', 'E2E family child');", dbFailOnError
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, IDExpedientePadre, OrdinalE2E, Nemotecnico, Titulo) VALUES (" & FamilySiblingId() & ", " & FamilyRootId() & ", 9303, 'E2E-FAMILY-SIBLING', 'E2E family sibling');", dbFailOnError
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, IDExpedientePadre, OrdinalE2E, Nemotecnico, Titulo) VALUES (" & FamilyGrandchildId() & ", " & FamilyChildId() & ", 9304, 'E2E-FAMILY-GRANDCHILD', 'E2E family grandchild');", dbFailOnError

    SeedE2EFamilyExportFixture = True
    Exit Function

EH:
    p_Error = "SeedE2EFamilyExportFixture: " & Err.Description
End Function

Private Function SeedE2EFamilyExportCycleFixture(ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim db As DAO.Database
    p_Error = ""
    SeedE2EFamilyExportCycleFixture = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function
    Set db = getdb(p_Error)
    If db Is Nothing Then
        If p_Error = "" Then p_Error = "SeedE2EFamilyExportCycleFixture: getdb returned Nothing"
        Exit Function
    End If

    db.Execute "INSERT INTO TbExpedientes (IDExpediente, IDExpedientePadre, OrdinalE2E, Nemotecnico, Titulo) VALUES (" & CycleFirstId() & ", Null, 9310, 'E2E-FAMILY-CYCLE-A', 'E2E family cycle A');", dbFailOnError
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, IDExpedientePadre, OrdinalE2E, Nemotecnico, Titulo) VALUES (" & CycleSecondId() & ", " & CycleFirstId() & ", 9311, 'E2E-FAMILY-CYCLE-B', 'E2E family cycle B');", dbFailOnError
    db.Execute "UPDATE TbExpedientes SET IDExpedientePadre=" & CycleSecondId() & " WHERE IDExpediente=" & CycleFirstId() & ";", dbFailOnError

    SeedE2EFamilyExportCycleFixture = True
    Exit Function

EH:
    p_Error = "SeedE2EFamilyExportCycleFixture: " & Err.Description
End Function

Private Function TeardownE2EFamilyExportFixture(ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim db As DAO.Database
    p_Error = ""
    TeardownE2EFamilyExportFixture = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function
    Set db = getdb(p_Error)
    If db Is Nothing Then
        If p_Error = "" Then p_Error = "TeardownE2EFamilyExportFixture: getdb returned Nothing"
        Exit Function
    End If

    db.Execute "UPDATE TbExpedientes SET IDExpedientePadre=Null WHERE IDExpediente BETWEEN " & FAMILY_EXPORT_TEST_ID_BASE & " AND " & (FAMILY_EXPORT_TEST_ID_BASE + 99) & ";", dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente BETWEEN " & FAMILY_EXPORT_TEST_ID_BASE & " AND " & (FAMILY_EXPORT_TEST_ID_BASE + 99) & ";", dbFailOnError

    TeardownE2EFamilyExportFixture = True
    Exit Function

EH:
    p_Error = "TeardownE2EFamilyExportFixture: " & Err.Description
End Function

Private Function CountIdInCanonicalJson(ByVal p_Json As String, ByVal p_IDExpediente As Long) As Long
    Dim token As String
    Dim position As Long

    token = Chr$(34) & "idexpediente" & Chr$(34) & ":" & CStr(p_IDExpediente)
    position = InStr(1, p_Json, token, vbTextCompare)
    Do While position > 0
        CountIdInCanonicalJson = CountIdInCanonicalJson + 1
        position = InStr(position + Len(token), p_Json, token, vbTextCompare)
    Loop

    token = Chr$(34) & "idexpediente" & Chr$(34) & ": " & CStr(p_IDExpediente)
    position = InStr(1, p_Json, token, vbTextCompare)
    Do While position > 0
        CountIdInCanonicalJson = CountIdInCanonicalJson + 1
        position = InStr(position + Len(token), p_Json, token, vbTextCompare)
    Loop
End Function

Private Function JsonContainsId(ByVal p_Json As String, ByVal p_IDExpediente As Long) As Boolean
    JsonContainsId = (CountIdInCanonicalJson(p_Json, p_IDExpediente) > 0)
End Function

Private Function BuildFamilyExportResult(ByVal p_Ok As Boolean, ByVal p_Value As String, ByRef p_Logs() As String) As String
    If p_Ok Then
        BuildFamilyExportResult = BuildJsonOk(p_Value, p_Logs)
    Else
        BuildFamilyExportResult = BuildJsonFail(p_Value, p_Logs)
    End If
End Function

Private Function NullOrdinalFirstId() As Long
    NullOrdinalFirstId = FAMILY_EXPORT_TEST_ID_BASE + 20
End Function

Private Function NullOrdinalSecondId() As Long
    NullOrdinalSecondId = FAMILY_EXPORT_TEST_ID_BASE + 21
End Function

Private Function SeedE2EFamilyExportNullOrdinalFixture(ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim db As DAO.Database
    p_Error = ""
    SeedE2EFamilyExportNullOrdinalFixture = False

    If Not EnsureSandboxBackend(p_Error) Then Exit Function
    Set db = getdb(p_Error)
    If db Is Nothing Then
        If p_Error = "" Then p_Error = "SeedE2EFamilyExportNullOrdinalFixture: getdb returned Nothing"
        Exit Function
    End If

    db.Execute "INSERT INTO TbExpedientes (IDExpediente, IDExpedientePadre, OrdinalE2E, Nemotecnico, Titulo) VALUES (" & NullOrdinalFirstId() & ", Null, Null, 'E2E-FAMILY-NULL-ORD-1', 'E2E family null ordinal 1');", dbFailOnError
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, IDExpedientePadre, OrdinalE2E, Nemotecnico, Titulo) VALUES (" & NullOrdinalSecondId() & ", " & NullOrdinalFirstId() & ", Null, 'E2E-FAMILY-NULL-ORD-2', 'E2E family null ordinal 2');", dbFailOnError

    SeedE2EFamilyExportNullOrdinalFixture = True
    Exit Function

EH:
    p_Error = "SeedE2EFamilyExportNullOrdinalFixture: " & Err.Description
End Function

Private Function ReadMaxOrdinalE2E(ByVal p_Db As DAO.Database) As Long
    Dim rs As DAO.Recordset

    Set rs = p_Db.OpenRecordset("SELECT Max(OrdinalE2E) AS MaxOrdinal FROM TbExpedientes WHERE OrdinalE2E IS NOT NULL;", dbOpenSnapshot)
    If Not rs.EOF Then ReadMaxOrdinalE2E = CLng(Nz(rs.Fields("MaxOrdinal").value, 0))
    rs.Close
    Set rs = Nothing
End Function

Private Function ReadOrdinalE2EValue(ByVal p_Db As DAO.Database, ByVal p_IDExpediente As Long) As Variant
    Dim rs As DAO.Recordset

    Set rs = p_Db.OpenRecordset("SELECT OrdinalE2E FROM TbExpedientes WHERE IDExpediente=" & p_IDExpediente & ";", dbOpenSnapshot)
    If Not rs.EOF Then ReadOrdinalE2EValue = rs.Fields("OrdinalE2E").value
    rs.Close
    Set rs = Nothing
End Function

Private Function JsonContainsExactToken(ByVal p_Json As String, ByVal p_Token As String) As Boolean
    JsonContainsExactToken = (InStr(1, p_Json, p_Token, vbBinaryCompare) > 0)
End Function

Private Function JsonContainsNumericToken(ByVal p_Json As String, ByVal p_Key As String, ByVal p_Value As Long) As Boolean
    JsonContainsNumericToken = JsonContainsExactToken(p_Json, Chr$(34) & p_Key & Chr$(34) & ":" & CStr(p_Value))
    If Not JsonContainsNumericToken Then
        JsonContainsNumericToken = JsonContainsExactToken(p_Json, Chr$(34) & p_Key & Chr$(34) & ": " & CStr(p_Value))
    End If
End Function

Public Function Test_E2EFamilyExport_EmptyInputReportsError() As String
    Dim errLocal As String
    Dim jsonText As String
    Dim logs(0 To 3) As String

    Test_E2EFamilyExport_EmptyInputReportsError = BuildJsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "Act: call family-aware canonical export with empty CSV"
    jsonText = GenerarJsonE2ECanonicoPorListaConFamilia("   ", errLocal)

    logs(1) = "Assert: empty input uses ByRef error channel and returns no JSON payload"
    If errLocal <> "" And Trim$(jsonText) = "" Then
        Test_E2EFamilyExport_EmptyInputReportsError = BuildFamilyExportResult(True, "empty input reports error", logs)
    Else
        Test_E2EFamilyExport_EmptyInputReportsError = BuildFamilyExportResult(False, "expected ByRef error and empty JSON for empty input", logs)
    End If
    Exit Function

HandleError:
    logs(2) = "Error: " & Err.Description
    Test_E2EFamilyExport_EmptyInputReportsError = BuildFamilyExportResult(False, "Test_E2EFamilyExport_EmptyInputReportsError: " & Err.Description, logs)
End Function

Public Function Test_E2EFamilyExport_ChildExpandsRootSiblingsAndDescendants() As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim jsonText As String
    Dim logs(0 To 6) As String

    Test_E2EFamilyExport_ChildExpandsRootSiblingsAndDescendants = BuildJsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "Arrange: setup sandbox and seed root, selected child, sibling, and grandchild"
    If Not SetupE2EFamilyExportSandbox(errLocal) Or errLocal <> "" Then GoTo FixtureFailed
    If Not SeedE2EFamilyExportFixture(errLocal) Or errLocal <> "" Then GoTo FixtureFailed

    logs(1) = "Act: export from non-root child"
    jsonText = GenerarJsonE2ECanonicoPorListaConFamilia(CStr(FamilyChildId()), errLocal)
    If errLocal <> "" Then
        Test_E2EFamilyExport_ChildExpandsRootSiblingsAndDescendants = BuildFamilyExportResult(False, "family export failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(2) = "Assert: root, selected child, sibling, and descendant are present"
    If JsonContainsId(jsonText, FamilyRootId()) And JsonContainsId(jsonText, FamilyChildId()) And JsonContainsId(jsonText, FamilySiblingId()) And JsonContainsId(jsonText, FamilyGrandchildId()) Then
        Test_E2EFamilyExport_ChildExpandsRootSiblingsAndDescendants = BuildFamilyExportResult(True, "child expands to full family", logs)
    Else
        Test_E2EFamilyExport_ChildExpandsRootSiblingsAndDescendants = BuildFamilyExportResult(False, "expected root, child, sibling, and grandchild in family export", logs)
    End If
    GoTo Cleanup

FixtureFailed:
    Test_E2EFamilyExport_ChildExpandsRootSiblingsAndDescendants = BuildFamilyExportResult(False, "fixture failed: " & errLocal, logs)

Cleanup:
    cleanupErr = ""
    If Not TeardownE2EFamilyExportFixture(cleanupErr) Or cleanupErr <> "" Then
        logs(5) = "Cleanup failed: " & cleanupErr
        Test_E2EFamilyExport_ChildExpandsRootSiblingsAndDescendants = BuildFamilyExportResult(False, "cleanup failed: " & cleanupErr, logs)
    End If
    logs(6) = "Cleanup complete"
    Exit Function

HandleError:
    logs(4) = "Error: " & Err.Description
    Test_E2EFamilyExport_ChildExpandsRootSiblingsAndDescendants = BuildFamilyExportResult(False, "Test_E2EFamilyExport_ChildExpandsRootSiblingsAndDescendants: " & Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_E2EFamilyExport_OverlappingSelectionsDedupeIds() As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim jsonText As String
    Dim logs(0 To 6) As String

    Test_E2EFamilyExport_OverlappingSelectionsDedupeIds = BuildJsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "Arrange: setup sandbox and seed overlapping family selections"
    If Not SetupE2EFamilyExportSandbox(errLocal) Or errLocal <> "" Then GoTo FixtureFailed
    If Not SeedE2EFamilyExportFixture(errLocal) Or errLocal <> "" Then GoTo FixtureFailed

    logs(1) = "Act: export from root and descendant in the same family"
    jsonText = GenerarJsonE2ECanonicoPorListaConFamilia(CStr(FamilyRootId()) & "," & CStr(FamilyChildId()), errLocal)
    If errLocal <> "" Then
        Test_E2EFamilyExport_OverlappingSelectionsDedupeIds = BuildFamilyExportResult(False, "family export failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(2) = "Assert: each family ID appears exactly once"
    If CountIdInCanonicalJson(jsonText, FamilyRootId()) = 1 And CountIdInCanonicalJson(jsonText, FamilyChildId()) = 1 And CountIdInCanonicalJson(jsonText, FamilySiblingId()) = 1 And CountIdInCanonicalJson(jsonText, FamilyGrandchildId()) = 1 Then
        Test_E2EFamilyExport_OverlappingSelectionsDedupeIds = BuildFamilyExportResult(True, "overlapping selections dedupe IDs", logs)
    Else
        Test_E2EFamilyExport_OverlappingSelectionsDedupeIds = BuildFamilyExportResult(False, "expected each family ID exactly once", logs)
    End If
    GoTo Cleanup

FixtureFailed:
    Test_E2EFamilyExport_OverlappingSelectionsDedupeIds = BuildFamilyExportResult(False, "fixture failed: " & errLocal, logs)

Cleanup:
    cleanupErr = ""
    If Not TeardownE2EFamilyExportFixture(cleanupErr) Or cleanupErr <> "" Then
        logs(5) = "Cleanup failed: " & cleanupErr
        Test_E2EFamilyExport_OverlappingSelectionsDedupeIds = BuildFamilyExportResult(False, "cleanup failed: " & cleanupErr, logs)
    End If
    logs(6) = "Cleanup complete"
    Exit Function

HandleError:
    logs(4) = "Error: " & Err.Description
    Test_E2EFamilyExport_OverlappingSelectionsDedupeIds = BuildFamilyExportResult(False, "Test_E2EFamilyExport_OverlappingSelectionsDedupeIds: " & Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_E2EFamilyExport_CycleTerminatesWithIdsOrError() As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim jsonText As String
    Dim logs(0 To 6) As String

    Test_E2EFamilyExport_CycleTerminatesWithIdsOrError = BuildJsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "Arrange: setup sandbox and seed a two-node parent cycle"
    If Not SetupE2EFamilyExportSandbox(errLocal) Or errLocal <> "" Then GoTo FixtureFailed
    If Not SeedE2EFamilyExportCycleFixture(errLocal) Or errLocal <> "" Then GoTo FixtureFailed

    logs(1) = "Act: export from a cycle member"
    jsonText = GenerarJsonE2ECanonicoPorListaConFamilia(CStr(CycleFirstId()), errLocal)

    logs(2) = "Assert: traversal terminates with valid JSON and each cycle member exactly once"
    If errLocal <> "" Then
        Test_E2EFamilyExport_CycleTerminatesWithIdsOrError = BuildFamilyExportResult(False, "expected valid JSON for cycle fixture, got error: " & errLocal, logs)
    ElseIf CountIdInCanonicalJson(jsonText, CycleFirstId()) = 1 And CountIdInCanonicalJson(jsonText, CycleSecondId()) = 1 Then
        Test_E2EFamilyExport_CycleTerminatesWithIdsOrError = BuildFamilyExportResult(True, "cycle terminates with each visited ID once", logs)
    Else
        Test_E2EFamilyExport_CycleTerminatesWithIdsOrError = BuildFamilyExportResult(False, "expected exactly one JSON row per cycle ID", logs)
    End If
    GoTo Cleanup

FixtureFailed:
    Test_E2EFamilyExport_CycleTerminatesWithIdsOrError = BuildFamilyExportResult(False, "fixture failed: " & errLocal, logs)

Cleanup:
    cleanupErr = ""
    If Not TeardownE2EFamilyExportFixture(cleanupErr) Or cleanupErr <> "" Then
        logs(5) = "Cleanup failed: " & cleanupErr
        Test_E2EFamilyExport_CycleTerminatesWithIdsOrError = BuildFamilyExportResult(False, "cleanup failed: " & cleanupErr, logs)
    End If
    logs(6) = "Cleanup complete"
    Exit Function

HandleError:
    logs(4) = "Error: " & Err.Description
    Test_E2EFamilyExport_CycleTerminatesWithIdsOrError = BuildFamilyExportResult(False, "Test_E2EFamilyExport_CycleTerminatesWithIdsOrError: " & Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_E2EFamilyExport_ListOnlyCompatibilityDoesNotExpandFamily() As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim jsonText As String
    Dim logs(0 To 6) As String

    Test_E2EFamilyExport_ListOnlyCompatibilityDoesNotExpandFamily = BuildJsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "Arrange: setup sandbox and seed a family with descendants"
    If Not SetupE2EFamilyExportSandbox(errLocal) Or errLocal <> "" Then GoTo FixtureFailed
    If Not SeedE2EFamilyExportFixture(errLocal) Or errLocal <> "" Then GoTo FixtureFailed

    logs(1) = "Act: call existing list-only canonical export from child"
    jsonText = GenerarJsonE2ECanonicoPorLista(CStr(FamilyChildId()), errLocal)
    If errLocal <> "" Then
        Test_E2EFamilyExport_ListOnlyCompatibilityDoesNotExpandFamily = BuildFamilyExportResult(False, "list-only export failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(2) = "Assert: existing helper keeps only requested ID"
    If CountIdInCanonicalJson(jsonText, FamilyChildId()) = 1 And CountIdInCanonicalJson(jsonText, FamilyRootId()) = 0 And CountIdInCanonicalJson(jsonText, FamilySiblingId()) = 0 And CountIdInCanonicalJson(jsonText, FamilyGrandchildId()) = 0 Then
        Test_E2EFamilyExport_ListOnlyCompatibilityDoesNotExpandFamily = BuildFamilyExportResult(True, "list-only export remains list-only", logs)
    Else
        Test_E2EFamilyExport_ListOnlyCompatibilityDoesNotExpandFamily = BuildFamilyExportResult(False, "existing list export expanded family unexpectedly", logs)
    End If
    GoTo Cleanup

FixtureFailed:
    Test_E2EFamilyExport_ListOnlyCompatibilityDoesNotExpandFamily = BuildFamilyExportResult(False, "fixture failed: " & errLocal, logs)

Cleanup:
    cleanupErr = ""
    If Not TeardownE2EFamilyExportFixture(cleanupErr) Or cleanupErr <> "" Then
        logs(5) = "Cleanup failed: " & cleanupErr
        Test_E2EFamilyExport_ListOnlyCompatibilityDoesNotExpandFamily = BuildFamilyExportResult(False, "cleanup failed: " & cleanupErr, logs)
    End If
    logs(6) = "Cleanup complete"
    Exit Function

HandleError:
    logs(4) = "Error: " & Err.Description
    Test_E2EFamilyExport_ListOnlyCompatibilityDoesNotExpandFamily = BuildFamilyExportResult(False, "Test_E2EFamilyExport_ListOnlyCompatibilityDoesNotExpandFamily: " & Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_E2EFamilyExport_AssignsNullOrdinalsIncrementally() As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim jsonText As String
    Dim db As DAO.Database
    Dim maxBefore As Long
    Dim firstOrdinal As Variant
    Dim secondOrdinal As Variant
    Dim ordinalsAssigned As Boolean
    Dim logs(0 To 7) As String

    Test_E2EFamilyExport_AssignsNullOrdinalsIncrementally = BuildJsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "Arrange: setup sandbox and seed exported family members with NULL OrdinalE2E"
    If Not SetupE2EFamilyExportSandbox(errLocal) Or errLocal <> "" Then GoTo FixtureFailed
    If Not SeedE2EFamilyExportNullOrdinalFixture(errLocal) Or errLocal <> "" Then GoTo FixtureFailed
    Set db = getdb(errLocal)
    If db Is Nothing Or errLocal <> "" Then GoTo FixtureFailed
    maxBefore = ReadMaxOrdinalE2E(db)

    logs(1) = "Act: export family from a member whose OrdinalE2E is NULL"
    jsonText = GenerarJsonE2ECanonicoPorListaConFamilia(CStr(NullOrdinalFirstId()), errLocal)
    If errLocal <> "" Then
        Test_E2EFamilyExport_AssignsNullOrdinalsIncrementally = BuildFamilyExportResult(False, "family export failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(2) = "Assert: missing ordinals are persisted incrementally after previous max"
    firstOrdinal = ReadOrdinalE2EValue(db, NullOrdinalFirstId())
    secondOrdinal = ReadOrdinalE2EValue(db, NullOrdinalSecondId())
    If Not IsNull(firstOrdinal) Then
        If Not IsNull(secondOrdinal) Then
            ordinalsAssigned = (CLng(firstOrdinal) = maxBefore + 1 And CLng(secondOrdinal) = maxBefore + 2)
        End If
    End If
    If ordinalsAssigned And JsonContainsId(jsonText, NullOrdinalFirstId()) And JsonContainsId(jsonText, NullOrdinalSecondId()) Then
        Test_E2EFamilyExport_AssignsNullOrdinalsIncrementally = BuildFamilyExportResult(True, "NULL ordinals assigned incrementally", logs)
    Else
        Test_E2EFamilyExport_AssignsNullOrdinalsIncrementally = BuildFamilyExportResult(False, "expected NULL ordinals to become max+1 and max+2", logs)
    End If
    GoTo Cleanup

FixtureFailed:
    Test_E2EFamilyExport_AssignsNullOrdinalsIncrementally = BuildFamilyExportResult(False, "fixture failed: " & errLocal, logs)

Cleanup:
    cleanupErr = ""
    If Not TeardownE2EFamilyExportFixture(cleanupErr) Or cleanupErr <> "" Then
        logs(6) = "Cleanup failed: " & cleanupErr
        Test_E2EFamilyExport_AssignsNullOrdinalsIncrementally = BuildFamilyExportResult(False, "cleanup failed: " & cleanupErr, logs)
    End If
    logs(7) = "Cleanup complete"
    Exit Function

HandleError:
    logs(5) = "Error: " & Err.Description
    Test_E2EFamilyExport_AssignsNullOrdinalsIncrementally = BuildFamilyExportResult(False, "Test_E2EFamilyExport_AssignsNullOrdinalsIncrementally: " & Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_E2EFamilyExport_RollsBackOrdinalsOnPreparationFailure() As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim jsonText As String
    Dim db As DAO.Database
    Dim firstOrdinal As Variant
    Dim secondOrdinal As Variant
    Dim rollbackComplete As Boolean
    Dim logs(0 To 7) As String

    Test_E2EFamilyExport_RollsBackOrdinalsOnPreparationFailure = BuildJsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "Arrange: setup NULL ordinal fixture and enable the RED preparation-failure seam"
    If Not SetupE2EFamilyExportSandbox(errLocal) Or errLocal <> "" Then GoTo FixtureFailed
    If Not SeedE2EFamilyExportNullOrdinalFixture(errLocal) Or errLocal <> "" Then GoTo FixtureFailed
    Set db = getdb(errLocal)
    If db Is Nothing Or errLocal <> "" Then GoTo FixtureFailed
    E2EFamilyExport_SetOrdinalPreparationFailureForTest True

    logs(1) = "Act: force family export preparation to fail after ordinal assignment starts"
    jsonText = GenerarJsonE2ECanonicoPorListaConFamilia(CStr(NullOrdinalFirstId()), errLocal)

    logs(2) = "Assert: error uses ByRef channel and partial ordinal updates are rolled back"
    firstOrdinal = ReadOrdinalE2EValue(db, NullOrdinalFirstId())
    secondOrdinal = ReadOrdinalE2EValue(db, NullOrdinalSecondId())
    rollbackComplete = IsNull(firstOrdinal)
    If rollbackComplete Then rollbackComplete = IsNull(secondOrdinal)
    If errLocal <> "" And Trim$(jsonText) = "" And rollbackComplete Then
        Test_E2EFamilyExport_RollsBackOrdinalsOnPreparationFailure = BuildFamilyExportResult(True, "preparation failure rolls back ordinals", logs)
    Else
        Test_E2EFamilyExport_RollsBackOrdinalsOnPreparationFailure = BuildFamilyExportResult(False, "expected ByRef error, empty JSON, and NULL ordinals after rollback", logs)
    End If
    GoTo Cleanup

FixtureFailed:
    Test_E2EFamilyExport_RollsBackOrdinalsOnPreparationFailure = BuildFamilyExportResult(False, "fixture failed: " & errLocal, logs)

Cleanup:
    On Error Resume Next
    E2EFamilyExport_SetOrdinalPreparationFailureForTest False
    On Error GoTo 0
    cleanupErr = ""
    If Not TeardownE2EFamilyExportFixture(cleanupErr) Or cleanupErr <> "" Then
        logs(6) = "Cleanup failed: " & cleanupErr
        Test_E2EFamilyExport_RollsBackOrdinalsOnPreparationFailure = BuildFamilyExportResult(False, "cleanup failed: " & cleanupErr, logs)
    End If
    logs(7) = "Cleanup complete"
    Exit Function

HandleError:
    logs(5) = "Error: " & Err.Description
    Test_E2EFamilyExport_RollsBackOrdinalsOnPreparationFailure = BuildFamilyExportResult(False, "Test_E2EFamilyExport_RollsBackOrdinalsOnPreparationFailure: " & Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_E2EFamilyExport_CanonicalJsonUsesUppercaseOrdinalE2EOnly() As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim jsonText As String
    Dim logs(0 To 6) As String

    Test_E2EFamilyExport_CanonicalJsonUsesUppercaseOrdinalE2EOnly = BuildJsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "Arrange: setup sandbox and seed family fixture with non-null ordinals"
    If Not SetupE2EFamilyExportSandbox(errLocal) Or errLocal <> "" Then GoTo FixtureFailed
    If Not SeedE2EFamilyExportFixture(errLocal) Or errLocal <> "" Then GoTo FixtureFailed

    logs(1) = "Act: export family JSON"
    jsonText = GenerarJsonE2ECanonicoPorListaConFamilia(CStr(FamilyChildId()), errLocal)
    If errLocal <> "" Then
        Test_E2EFamilyExport_CanonicalJsonUsesUppercaseOrdinalE2EOnly = BuildFamilyExportResult(False, "family export failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(2) = "Assert: JSON uses canonical OrdinalE2E casing and no lowercase alias"
    If JsonContainsExactToken(jsonText, Chr$(34) & "OrdinalE2E" & Chr$(34)) And Not JsonContainsExactToken(jsonText, Chr$(34) & "ordinale2e" & Chr$(34)) Then
        Test_E2EFamilyExport_CanonicalJsonUsesUppercaseOrdinalE2EOnly = BuildFamilyExportResult(True, "canonical OrdinalE2E casing preserved", logs)
    Else
        Test_E2EFamilyExport_CanonicalJsonUsesUppercaseOrdinalE2EOnly = BuildFamilyExportResult(False, "expected uppercase OrdinalE2E and no lowercase alias", logs)
    End If
    GoTo Cleanup

FixtureFailed:
    Test_E2EFamilyExport_CanonicalJsonUsesUppercaseOrdinalE2EOnly = BuildFamilyExportResult(False, "fixture failed: " & errLocal, logs)

Cleanup:
    cleanupErr = ""
    If Not TeardownE2EFamilyExportFixture(cleanupErr) Or cleanupErr <> "" Then
        logs(5) = "Cleanup failed: " & cleanupErr
        Test_E2EFamilyExport_CanonicalJsonUsesUppercaseOrdinalE2EOnly = BuildFamilyExportResult(False, "cleanup failed: " & cleanupErr, logs)
    End If
    logs(6) = "Cleanup complete"
    Exit Function

HandleError:
    logs(4) = "Error: " & Err.Description
    Test_E2EFamilyExport_CanonicalJsonUsesUppercaseOrdinalE2EOnly = BuildFamilyExportResult(False, "Test_E2EFamilyExport_CanonicalJsonUsesUppercaseOrdinalE2EOnly: " & Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_E2EFamilyExport_BatchJsonDelegatesToFamilyExportCore() As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim jsonText As String
    Dim logs(0 To 6) As String

    Test_E2EFamilyExport_BatchJsonDelegatesToFamilyExportCore = BuildJsonFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "Arrange: setup sandbox and seed a family selected by batch export"
    If Not SetupE2EFamilyExportSandbox(errLocal) Or errLocal <> "" Then GoTo FixtureFailed
    If Not SeedE2EFamilyExportFixture(errLocal) Or errLocal <> "" Then GoTo FixtureFailed

    logs(1) = "Act: generate batch JSON from selected child"
    jsonText = GenerateE2EExportJsonForBatch(12345, CStr(FamilyChildId()), errLocal)
    If errLocal <> "" Then
        Test_E2EFamilyExport_BatchJsonDelegatesToFamilyExportCore = BuildFamilyExportResult(False, "batch export failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(2) = "Assert: batch JSON keeps meta and contains delegated family data"
    If JsonContainsNumericToken(jsonText, "IDExportacion", 12345) And JsonContainsNumericToken(jsonText, "IDBatch", 12345) And JsonContainsId(jsonText, FamilyRootId()) And JsonContainsId(jsonText, FamilyChildId()) And JsonContainsId(jsonText, FamilySiblingId()) And JsonContainsId(jsonText, FamilyGrandchildId()) Then
        Test_E2EFamilyExport_BatchJsonDelegatesToFamilyExportCore = BuildFamilyExportResult(True, "batch JSON delegates to family export core", logs)
    Else
        Test_E2EFamilyExport_BatchJsonDelegatesToFamilyExportCore = BuildFamilyExportResult(False, "expected batch meta plus root, child, sibling, and grandchild data", logs)
    End If
    GoTo Cleanup

FixtureFailed:
    Test_E2EFamilyExport_BatchJsonDelegatesToFamilyExportCore = BuildFamilyExportResult(False, "fixture failed: " & errLocal, logs)

Cleanup:
    cleanupErr = ""
    If Not TeardownE2EFamilyExportFixture(cleanupErr) Or cleanupErr <> "" Then
        logs(5) = "Cleanup failed: " & cleanupErr
        Test_E2EFamilyExport_BatchJsonDelegatesToFamilyExportCore = BuildFamilyExportResult(False, "cleanup failed: " & cleanupErr, logs)
    End If
    logs(6) = "Cleanup complete"
    Exit Function

HandleError:
    logs(4) = "Error: " & Err.Description
    Test_E2EFamilyExport_BatchJsonDelegatesToFamilyExportCore = BuildFamilyExportResult(False, "Test_E2EFamilyExport_BatchJsonDelegatesToFamilyExportCore: " & Err.Description, logs)
    Resume Cleanup
End Function
