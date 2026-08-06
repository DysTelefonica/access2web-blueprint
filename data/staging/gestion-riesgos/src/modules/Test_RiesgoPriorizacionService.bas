Attribute VB_Name = "Test_RiesgoPriorizacionService"
Option Compare Database
Option Explicit

Private Const FIX_RISK_ID_1 As Long = 936001
Private Const FIX_RISK_ID_2 As Long = 936002
Private Const FIX_RISK_ID_3 As Long = 936003

' Schema evidence captured through Dysflow MCP on frontend/local Gestion_Riesgos.accdb:
' TbAuxPriorizacion fields: IDRiesgo Long nullable; Pri, PriEdAnterior, Codigo,
' Descripcion, CausaRaiz Text(255) nullable/zero-length allowed. No user PK/FK
' relationships reference TbAuxPriorizacion in the frontend relationship catalog.

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

Private Function SqlText(ByVal value As String) As String
    SqlText = "'" & Replace(value, "'", "''") & "'"
End Function

Private Function PriorityMap( _
    ByVal pri1 As Variant, _
    ByVal pri2 As Variant, _
    ByVal pri3 As Variant _
) As Scripting.Dictionary
    Dim values As New Scripting.Dictionary
    values.CompareMode = TextCompare
    values.Add CStr(FIX_RISK_ID_1), pri1
    values.Add CStr(FIX_RISK_ID_2), pri2
    values.Add CStr(FIX_RISK_ID_3), pri3
    Set PriorityMap = values
End Function

Private Function AssertValidationCase( _
    ByVal caseName As String, _
    ByVal priorities As Scripting.Dictionary, _
    ByVal expectedOk As Boolean, _
    ByVal expectedErrorFragment As String, _
    ByRef logs() As String _
) As String
    On Error GoTo EH

    Dim service As RiesgoPriorizacionService
    Dim errMsg As String
    Dim ok As Boolean

    Set service = New RiesgoPriorizacionService
    ok = service.ValidatePriorities(priorities, 3, errMsg)

    If ok <> expectedOk Then
        AssertValidationCase = BuildFail(caseName & ": expected ok=" & CStr(expectedOk) & _
            ", got ok=" & CStr(ok) & ", error=" & errMsg, logs)
        Exit Function
    End If
    If expectedErrorFragment = "" Then
        If errMsg <> "" Then
            AssertValidationCase = BuildFail(caseName & ": expected empty error, got " & errMsg, logs)
            Exit Function
        End If
    ElseIf InStr(1, errMsg, expectedErrorFragment, vbTextCompare) = 0 Then
        AssertValidationCase = BuildFail(caseName & ": expected error containing '" & _
            expectedErrorFragment & "', got '" & errMsg & "'", logs)
        Exit Function
    End If

    AssertValidationCase = BuildOk(caseName, logs)
    Exit Function

EH:
    AssertValidationCase = BuildFail(caseName & ": " & Err.Description, logs)
End Function

Private Function AssertChangedRiskIds( _
    ByVal caseName As String, _
    ByVal initialPriorities As Scripting.Dictionary, _
    ByVal finalPriorities As Scripting.Dictionary, _
    ByVal expectedChangedIdsCsv As String, _
    ByVal expectedErrorFragment As String, _
    ByRef logs() As String _
) As String
    On Error GoTo EH

    Dim service As RiesgoPriorizacionService
    Dim changedRiskIds As Scripting.Dictionary
    Dim errMsg As String
    Dim expectedIds As Variant
    Dim expectedId As Variant

    Set service = New RiesgoPriorizacionService
    Set changedRiskIds = service.GetChangedRiskIds(initialPriorities, finalPriorities, errMsg)

    If expectedErrorFragment <> "" Then
        If errMsg = "" Then
            AssertChangedRiskIds = BuildFail(caseName & ": expected error containing '" & _
                expectedErrorFragment & "'", logs)
            Exit Function
        End If
        If InStr(1, errMsg, expectedErrorFragment, vbTextCompare) = 0 Then
            AssertChangedRiskIds = BuildFail(caseName & ": expected error containing '" & _
                expectedErrorFragment & "', got '" & errMsg & "'", logs)
            Exit Function
        End If
        AssertChangedRiskIds = BuildOk(caseName, logs)
        Exit Function
    End If

    If errMsg <> "" Then
        AssertChangedRiskIds = BuildFail(caseName & ": unexpected error: " & errMsg, logs)
        Exit Function
    End If
    If changedRiskIds Is Nothing Then
        AssertChangedRiskIds = BuildFail(caseName & ": changedRiskIds is Nothing", logs)
        Exit Function
    End If

    If expectedChangedIdsCsv = "" Then
        If changedRiskIds.Count <> 0 Then
            AssertChangedRiskIds = BuildFail(caseName & ": expected no changed risks, got " & _
                CStr(changedRiskIds.Count), logs)
            Exit Function
        End If
        AssertChangedRiskIds = BuildOk(caseName, logs)
        Exit Function
    End If

    expectedIds = Split(expectedChangedIdsCsv, ",")
    If changedRiskIds.Count <> UBound(expectedIds) - LBound(expectedIds) + 1 Then
        AssertChangedRiskIds = BuildFail(caseName & ": expected changed count " & _
            CStr(UBound(expectedIds) - LBound(expectedIds) + 1) & ", got " & _
            CStr(changedRiskIds.Count), logs)
        Exit Function
    End If

    For Each expectedId In expectedIds
        If Not changedRiskIds.Exists(CStr(expectedId)) Then
            AssertChangedRiskIds = BuildFail(caseName & ": missing changed risk id " & _
                CStr(expectedId), logs)
            Exit Function
        End If
        If CStr(changedRiskIds(CStr(expectedId))) <> CStr(finalPriorities(CStr(expectedId))) Then
            AssertChangedRiskIds = BuildFail(caseName & ": changed priority mismatch for " & _
                CStr(expectedId), logs)
            Exit Function
        End If
    Next expectedId

    AssertChangedRiskIds = BuildOk(caseName, logs)
    Exit Function

EH:
    AssertChangedRiskIds = BuildFail(caseName & ": " & Err.Description, logs)
End Function

Private Sub DeleteIssue36Rows(ByVal db As DAO.Database)
    db.Execute "DELETE FROM TbAuxPriorizacion WHERE IDRiesgo IN (" & _
        CStr(FIX_RISK_ID_1) & "," & CStr(FIX_RISK_ID_2) & "," & CStr(FIX_RISK_ID_3) & ")", _
        dbFailOnError
End Sub

Private Function CountIssue36Rows(ByVal db As DAO.Database) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS Cnt FROM TbAuxPriorizacion WHERE IDRiesgo IN (" & _
        CStr(FIX_RISK_ID_1) & "," & CStr(FIX_RISK_ID_2) & "," & CStr(FIX_RISK_ID_3) & ")", _
        dbOpenSnapshot)
    CountIssue36Rows = CLng(Nz(rs.Fields("Cnt").Value, 0))
    rs.Close
    Set rs = Nothing
End Function

Public Function Test_RiesgoPriorizacionService_Validate_AcceptsCompleteUniqueRange() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: priorities 1, 2, 3 for deterministic risk IDs"
    logs(1) = "2. Act: ValidatePriorities"
    logs(2) = "3. Assert: validation succeeds with empty error"
    Test_RiesgoPriorizacionService_Validate_AcceptsCompleteUniqueRange = _
        AssertValidationCase("valid complete range", PriorityMap("1", "2", "3"), True, "", logs)
End Function

Public Function Test_RiesgoPriorizacionService_Validate_RejectsNonNumeric() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: one priority is non-numeric ABC"
    logs(1) = "2. Act: ValidatePriorities"
    logs(2) = "3. Assert: validation fails as non-numeric"
    Test_RiesgoPriorizacionService_Validate_RejectsNonNumeric = _
        AssertValidationCase("non numeric", PriorityMap("1", "ABC", "3"), False, "número", logs)
End Function

Public Function Test_RiesgoPriorizacionService_Validate_RejectsLowerThanOne() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: one priority is 0"
    logs(1) = "2. Act: ValidatePriorities"
    logs(2) = "3. Assert: validation fails as outside 1..N"
    Test_RiesgoPriorizacionService_Validate_RejectsLowerThanOne = _
        AssertValidationCase("lower than one", PriorityMap("0", "2", "3"), False, "entre el 1 y el 3", logs)
End Function

Public Function Test_RiesgoPriorizacionService_Validate_RejectsGreaterThanCount() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: one priority is 4 while risk count is 3"
    logs(1) = "2. Act: ValidatePriorities"
    logs(2) = "3. Assert: validation fails as outside 1..3"
    Test_RiesgoPriorizacionService_Validate_RejectsGreaterThanCount = _
        AssertValidationCase("greater than count", PriorityMap("1", "2", "4"), False, "entre el 1 y el 3", logs)
End Function

Public Function Test_RiesgoPriorizacionService_Validate_RejectsDuplicate() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: priorities contain duplicate 2"
    logs(1) = "2. Act: ValidatePriorities"
    logs(2) = "3. Assert: validation fails identifying duplicate 2"
    Test_RiesgoPriorizacionService_Validate_RejectsDuplicate = _
        AssertValidationCase("duplicate priority", PriorityMap("1", "2", "2"), False, "2", logs)
End Function

Public Function Test_RiesgoPriorizacionService_Validate_RejectsMissingPriority() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: one priority is empty"
    logs(1) = "2. Act: ValidatePriorities"
    logs(2) = "3. Assert: validation fails as incomplete"
    Test_RiesgoPriorizacionService_Validate_RejectsMissingPriority = _
        AssertValidationCase("missing priority", PriorityMap("1", "", "3"), False, "incluir", logs)
End Function

Public Function Test_RiesgoPriorizacionService_Validate_RejectsDecimalPriority() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: one priority is decimal 1.5"
    logs(1) = "2. Act: ValidatePriorities"
    logs(2) = "3. Assert: validation fails because priorities must be whole numbers"
    Test_RiesgoPriorizacionService_Validate_RejectsDecimalPriority = _
        AssertValidationCase("decimal priority", PriorityMap("1", "1.5", "3"), False, "número", logs)
End Function

Public Function Test_RiesgoPriorizacionService_Changes_NoChanges_ReturnsEmpty() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: initial and final priorities are identical"
    logs(1) = "2. Act: GetChangedRiskIds"
    logs(2) = "3. Assert: returned dictionary is empty"
    Test_RiesgoPriorizacionService_Changes_NoChanges_ReturnsEmpty = _
        AssertChangedRiskIds("no priority changes", PriorityMap("1", "2", "3"), _
            PriorityMap("1", "2", "3"), "", "", logs)
End Function

Public Function Test_RiesgoPriorizacionService_Changes_OneChanged_ReturnsRiskId() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: one final priority differs from the initial snapshot"
    logs(1) = "2. Act: GetChangedRiskIds"
    logs(2) = "3. Assert: only the changed risk ID is returned"
    Test_RiesgoPriorizacionService_Changes_OneChanged_ReturnsRiskId = _
        AssertChangedRiskIds("one priority changed", PriorityMap("1", "2", "3"), _
            PriorityMap("1", "1", "3"), CStr(FIX_RISK_ID_2), "", logs)
End Function

Public Function Test_RiesgoPriorizacionService_Changes_MultipleChanged_ReturnsRiskIds() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: multiple final priorities differ from the initial snapshot"
    logs(1) = "2. Act: GetChangedRiskIds"
    logs(2) = "3. Assert: all changed risk IDs are returned"
    Test_RiesgoPriorizacionService_Changes_MultipleChanged_ReturnsRiskIds = _
        AssertChangedRiskIds("multiple priorities changed", PriorityMap("1", "2", "3"), _
            PriorityMap("3", "2", "1"), CStr(FIX_RISK_ID_1) & "," & CStr(FIX_RISK_ID_3), "", logs)
End Function

Public Function Test_RiesgoPriorizacionService_Changes_WhitespaceDifference_ReturnsRiskId() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: final priority differs only by leading whitespace"
    logs(1) = "2. Act: GetChangedRiskIds"
    logs(2) = "3. Assert: direct string comparison treats whitespace as a change"
    Test_RiesgoPriorizacionService_Changes_WhitespaceDifference_ReturnsRiskId = _
        AssertChangedRiskIds("whitespace-sensitive priority change", PriorityMap("1", "2", "3"), _
            PriorityMap(" 1", "2", "3"), CStr(FIX_RISK_ID_1), "", logs)
End Function

Public Function Test_RiesgoPriorizacionService_Changes_NothingInitial_ReturnsError() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: initial priorities are Nothing"
    logs(1) = "2. Act: GetChangedRiskIds"
    logs(2) = "3. Assert: error is returned without UI or DB access"
    Test_RiesgoPriorizacionService_Changes_NothingInitial_ReturnsError = _
        AssertChangedRiskIds("nothing initial", Nothing, PriorityMap("1", "2", "3"), _
            "", "listado inicial", logs)
End Function

Public Function Test_RiesgoPriorizacionService_Changes_NothingFinal_ReturnsError() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: final priorities are Nothing"
    logs(1) = "2. Act: GetChangedRiskIds"
    logs(2) = "3. Assert: error is returned without UI or DB access"
    Test_RiesgoPriorizacionService_Changes_NothingFinal_ReturnsError = _
        AssertChangedRiskIds("nothing final", PriorityMap("1", "2", "3"), Nothing, _
            "", "listado final", logs)
End Function

Public Function Test_RiesgoPriorizacionService_Changes_MissingInitialKey_ReturnsError() As String
    Dim logs(0 To 2) As String
    Dim initialPriorities As Scripting.Dictionary
    Dim finalPriorities As Scripting.Dictionary

    logs(0) = "1. Arrange: final priorities contain a risk absent from initial snapshot"
    logs(1) = "2. Act: GetChangedRiskIds"
    logs(2) = "3. Assert: error identifies the missing initial priority"

    Set initialPriorities = PriorityMap("1", "2", "3")
    initialPriorities.Remove CStr(FIX_RISK_ID_3)
    Set finalPriorities = PriorityMap("1", "2", "1")

    Test_RiesgoPriorizacionService_Changes_MissingInitialKey_ReturnsError = _
        AssertChangedRiskIds("missing initial key", initialPriorities, finalPriorities, _
            "", CStr(FIX_RISK_ID_3), logs)
End Function

Public Function Test_PriorizacionRepository_RoundTrip_UsesOnlyIssue36Fixtures() As String
    On Error GoTo EH

    Dim logs(0 To 7) As String
    Dim db As DAO.Database
    Dim priorities As Scripting.Dictionary
    Dim errMsg As String
    Dim rowsBefore As Long
    Dim rowsAfter As Long

    logs(0) = "1. Arrange: CurrentDb local/frontend TbAuxPriorizacion"
    logs(1) = "2. Arrange: delete only deterministic issue #36 IDs 936001..936003"
    logs(2) = "3. Assert pre-state: fixture cardinality is 0"
    logs(3) = "4. Act: insert three repository snapshot rows"
    logs(4) = "5. Act: update one temporary priority"
    logs(5) = "6. Act: read repository priorities"
    logs(6) = "7. Assert exact values and fixture cardinality"
    logs(7) = "8. Teardown: delete only issue #36 IDs"

    Set db = CurrentDb
    DeleteIssue36Rows db
    rowsBefore = CountIssue36Rows(db)
    If rowsBefore <> 0 Then
        Test_PriorizacionRepository_RoundTrip_UsesOnlyIssue36Fixtures = _
            BuildFail("Fixture pre-state expected 0 rows, got " & rowsBefore, logs)
        GoTo Teardown
    End If

    If Not Priorizacion_InsertSnapshotRow(FIX_RISK_ID_1, "1", "3", "RG-936001", "Issue 36 fixture 1", "", errMsg) Then
        Test_PriorizacionRepository_RoundTrip_UsesOnlyIssue36Fixtures = BuildFail(errMsg, logs)
        GoTo Teardown
    End If
    If Not Priorizacion_InsertSnapshotRow(FIX_RISK_ID_2, "2", "2", "RG-936002", "Issue 36 fixture 2", "", errMsg) Then
        Test_PriorizacionRepository_RoundTrip_UsesOnlyIssue36Fixtures = BuildFail(errMsg, logs)
        GoTo Teardown
    End If
    If Not Priorizacion_InsertSnapshotRow(FIX_RISK_ID_3, "3", "1", "RG-936003", "Issue 36 fixture 3", "", errMsg) Then
        Test_PriorizacionRepository_RoundTrip_UsesOnlyIssue36Fixtures = BuildFail(errMsg, logs)
        GoTo Teardown
    End If
    If Not Priorizacion_UpdateTempPriority(FIX_RISK_ID_2, "1", errMsg) Then
        Test_PriorizacionRepository_RoundTrip_UsesOnlyIssue36Fixtures = BuildFail(errMsg, logs)
        GoTo Teardown
    End If

    Set priorities = Priorizacion_ReadTempPriorities(errMsg)
    If priorities Is Nothing Then
        Test_PriorizacionRepository_RoundTrip_UsesOnlyIssue36Fixtures = BuildFail(errMsg, logs)
        GoTo Teardown
    End If

    rowsAfter = CountIssue36Rows(db)
    If rowsAfter <> 3 Then
        Test_PriorizacionRepository_RoundTrip_UsesOnlyIssue36Fixtures = _
            BuildFail("Fixture post-state expected 3 rows, got " & rowsAfter, logs)
        GoTo Teardown
    End If
    If CStr(priorities(CStr(FIX_RISK_ID_2))) <> "1" Then
        Test_PriorizacionRepository_RoundTrip_UsesOnlyIssue36Fixtures = _
            BuildFail("Updated priority for fixture 936002 should be 1", logs)
        GoTo Teardown
    End If

    Test_PriorizacionRepository_RoundTrip_UsesOnlyIssue36Fixtures = BuildOk("repository_roundtrip_issue36", logs)

Teardown:
    On Error Resume Next
    DeleteIssue36Rows db
    Set priorities = Nothing
    Set db = Nothing
    On Error GoTo 0
    Exit Function

EH:
    Test_PriorizacionRepository_RoundTrip_UsesOnlyIssue36Fixtures = _
        BuildFail("Test_PriorizacionRepository_RoundTrip_UsesOnlyIssue36Fixtures: " & Err.Description, logs)
    Resume Teardown
End Function
