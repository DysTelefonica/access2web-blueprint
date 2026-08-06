Attribute VB_Name = "Test_E2EBatchSchema"
Option Compare Database
Option Explicit

Public Function Test_E2EBatchSchema_EnsureCreatesRequiredBackendTables() As String
    Dim p_Error As String
    Dim db As DAO.Database
    Dim errLocal As String
    Dim cleanupErr As String
    Dim cleanupNeeded As Boolean
    Dim logs(0 To 4) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_E2EBatchSchema_EnsureCreatesRequiredBackendTables = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError
    p_Error = ""
    errLocal = ""

    logs(0) = "Arrange: setup sandbox and ensure E2E batch management backend schema"
    If Not SetupE2EBatchSchemaSandbox(errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EBatchSchema_EnsureCreatesRequiredBackendTables = JsonOK(False, "schema ensure failed: " & errLocal, logs)
        GoTo Cleanup
    End If
    cleanupNeeded = True

    If Not EnsureE2EBatchManagementSchema(errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EBatchSchema_EnsureCreatesRequiredBackendTables = JsonOK(False, "schema ensure failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    Set db = getdb(errLocal)
    If errLocal <> "" Or db Is Nothing Then
        p_Error = errLocal
        Test_E2EBatchSchema_EnsureCreatesRequiredBackendTables = JsonOK(False, "getdb failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(1) = "Assert: TbE2EExportSeleccionTemp has session selection contract"
    If Not AssertField(db, "TbE2EExportSeleccionTemp", "IDTemp", dbLong, False, logs) Then GoTo Fail
    If Not AssertPrimaryKey(db, "TbE2EExportSeleccionTemp", "IDTemp", logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportSeleccionTemp", "UsuarioConectado", dbText, True, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportSeleccionTemp", "SessionId", dbText, True, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportSeleccionTemp", "IDExpediente", dbLong, True, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportSeleccionTemp", "CreatedAt", dbDate, False, logs) Then GoTo Fail
    If Not AssertIndex(db, "TbE2EExportSeleccionTemp", "IX_TbE2EExportSeleccionTemp_Session", True, logs) Then GoTo Fail

    logs(2) = "Assert: persistent batch header exists for traceability"
    If Not AssertField(db, "TbE2EExportBatch", "IDBatch", dbLong, False, logs) Then GoTo Fail
    If Not AssertPrimaryKey(db, "TbE2EExportBatch", "IDBatch", logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportBatch", "SessionId", dbText, True, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportBatch", "UsuarioConectado", dbText, True, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportBatch", "Estado", dbText, True, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportBatch", "CreatedAt", dbDate, False, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportBatch", "StartedAt", dbDate, False, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportBatch", "CompletedAt", dbDate, False, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportBatch", "TotalSeleccionados", dbLong, False, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportBatch", "TotalExportados", dbLong, False, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportBatch", "ErrorMessage", dbMemo, False, logs) Then GoTo Fail
    If Not AssertIndex(db, "TbE2EExportBatch", "IX_TbE2EExportBatch_Session", False, logs) Then GoTo Fail

    logs(3) = "Assert: persistent batch detail exists for selected expediente traceability"
    If Not AssertField(db, "TbE2EExportBatchDetalle", "IDBatchDetalle", dbLong, False, logs) Then GoTo Fail
    If Not AssertPrimaryKey(db, "TbE2EExportBatchDetalle", "IDBatchDetalle", logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportBatchDetalle", "IDBatch", dbLong, True, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportBatchDetalle", "IDExpediente", dbLong, True, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportBatchDetalle", "OrdinalSeleccion", dbLong, False, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportBatchDetalle", "HashExportado", dbText, False, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportBatchDetalle", "Estado", dbText, True, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportBatchDetalle", "CreatedAt", dbDate, False, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EExportBatchDetalle", "ExportedAt", dbDate, False, logs) Then GoTo Fail
    If Not AssertIndex(db, "TbE2EExportBatchDetalle", "IX_TbE2EExportBatchDetalle_BatchExpediente", True, logs) Then GoTo Fail

    Test_E2EBatchSchema_EnsureCreatesRequiredBackendTables = JsonOK(True, "E2E batch schema contract is present", logs)
    GoTo Cleanup

Fail:
    Test_E2EBatchSchema_EnsureCreatesRequiredBackendTables = JsonOK(False, "E2E batch schema contract mismatch", logs)
    GoTo Cleanup

HandleError:
    p_Error = "Test_E2EBatchSchema_EnsureCreatesRequiredBackendTables: " & Err.Description
    logs(0) = p_Error
    Test_E2EBatchSchema_EnsureCreatesRequiredBackendTables = JsonOK(False, p_Error, logs)

Cleanup:
    On Error Resume Next
    If cleanupNeeded Then
        cleanupErr = ""
        If Not TeardownE2EBatchSchemaSandbox(cleanupErr) Or cleanupErr <> "" Then
            p_Error = cleanupErr
            Test_E2EBatchSchema_EnsureCreatesRequiredBackendTables = JsonOK(False, "cleanup failed: " & cleanupErr, logs)
        End If
    End If
End Function

Public Function Test_E2EBatchSchema_EnsureIsIdempotent() As String
    Dim p_Error As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim cleanupNeeded As Boolean
    Dim logs(0 To 1) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_E2EBatchSchema_EnsureIsIdempotent = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    p_Error = ""
    logs(0) = "Arrange: setup schema sandbox and run ensure twice"

    If Not SetupE2EBatchSchemaSandbox(errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EBatchSchema_EnsureIsIdempotent = JsonOK(False, "setup failed: " & errLocal, logs)
        GoTo Cleanup
    End If
    cleanupNeeded = True

    If Not EnsureE2EBatchManagementSchema(errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EBatchSchema_EnsureIsIdempotent = JsonOK(False, "first ensure failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    errLocal = ""
    If Not EnsureE2EBatchManagementSchema(errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EBatchSchema_EnsureIsIdempotent = JsonOK(False, "second ensure failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(1) = "Assert: second ensure completed without duplicate table/index errors"
    Test_E2EBatchSchema_EnsureIsIdempotent = JsonOK(True, "schema ensure is idempotent", logs)
    GoTo Cleanup

HandleError:
    p_Error = "Test_E2EBatchSchema_EnsureIsIdempotent: " & Err.Description
    logs(0) = p_Error
    Test_E2EBatchSchema_EnsureIsIdempotent = JsonOK(False, p_Error, logs)

Cleanup:
    On Error Resume Next
    If cleanupNeeded Then
        cleanupErr = ""
        If Not TeardownE2EBatchSchemaSandbox(cleanupErr) Or cleanupErr <> "" Then
            p_Error = cleanupErr
            Test_E2EBatchSchema_EnsureIsIdempotent = JsonOK(False, "cleanup failed: " & cleanupErr, logs)
        End If
    End If
End Function

Public Function Test_E2EBatchSchema_EnsureDestinationUserConfigContract() As String
    Dim p_Error As String
    Dim db As DAO.Database
    Dim errLocal As String
    Dim cleanupErr As String
    Dim cleanupNeeded As Boolean
    Dim logs(0 To 4) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_E2EBatchSchema_EnsureDestinationUserConfigContract = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError
    p_Error = ""
    errLocal = ""

    logs(0) = "Arrange: setup sandbox and ensure per-user destination schema"
    If Not SetupE2EBatchSchemaSandbox(errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EBatchSchema_EnsureDestinationUserConfigContract = JsonOK(False, "setup failed: " & errLocal, logs)
        GoTo Cleanup
    End If
    cleanupNeeded = True

    If Not EnsureE2EJsonDestinationConfigSchema(errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EBatchSchema_EnsureDestinationUserConfigContract = JsonOK(False, "schema ensure failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    Set db = getdb(errLocal)
    If errLocal <> "" Or db Is Nothing Then
        p_Error = errLocal
        Test_E2EBatchSchema_EnsureDestinationUserConfigContract = JsonOK(False, "getdb failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(1) = "Assert: destination table fields and PK"
    If Not AssertField(db, "TbE2EJsonDestinationUserConfig", "IDConfig", dbLong, False, logs) Then GoTo Fail
    If Not AssertPrimaryKey(db, "TbE2EJsonDestinationUserConfig", "IDConfig", logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EJsonDestinationUserConfig", "UsuarioRed", dbText, True, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EJsonDestinationUserConfig", "RutaDestino", dbMemo, True, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EJsonDestinationUserConfig", "CreatedAt", dbDate, False, logs) Then GoTo Fail
    If Not AssertField(db, "TbE2EJsonDestinationUserConfig", "UpdatedAt", dbDate, False, logs) Then GoTo Fail

    logs(2) = "Assert: unique index by UsuarioRed"
    If Not AssertIndex(db, "TbE2EJsonDestinationUserConfig", "UX_TbE2EJsonDestinationUserConfig_UsuarioRed", True, logs) Then GoTo Fail

    logs(3) = "Assert: unique index contains exactly UsuarioRed field"
    If Not AssertIndexFields(db, "TbE2EJsonDestinationUserConfig", "UX_TbE2EJsonDestinationUserConfig_UsuarioRed", "UsuarioRed", True, logs) Then GoTo Fail

    logs(4) = "Schema facts inspected: TbE2EJsonDestinationUserConfig has no DAO relationships/FKs; fixture cleanup is scoped by UsuarioRed LIKE 'qa.user%'"
    If Not AssertNoRelationshipsForTable(db, "TbE2EJsonDestinationUserConfig", logs) Then GoTo Fail

    Test_E2EBatchSchema_EnsureDestinationUserConfigContract = JsonOK(True, "destination user config contract is present", logs)
    GoTo Cleanup

Fail:
    Test_E2EBatchSchema_EnsureDestinationUserConfigContract = JsonOK(False, "destination user config contract mismatch", logs)
    GoTo Cleanup

HandleError:
    p_Error = "Test_E2EBatchSchema_EnsureDestinationUserConfigContract: " & Err.Description
    logs(0) = p_Error
    Test_E2EBatchSchema_EnsureDestinationUserConfigContract = JsonOK(False, p_Error, logs)

Cleanup:
    On Error Resume Next
    If cleanupNeeded Then
        cleanupErr = ""
        If Not TeardownE2EJsonDestinationConfigSandbox(cleanupErr) Or cleanupErr <> "" Then
            p_Error = cleanupErr
            Test_E2EBatchSchema_EnsureDestinationUserConfigContract = JsonOK(False, "cleanup failed: " & cleanupErr, logs)
        End If
    End If
End Function

Private Function AssertField( _
    ByVal p_Db As DAO.Database, _
    ByVal p_TableName As String, _
    ByVal p_FieldName As String, _
    ByVal p_Type As Integer, _
    ByVal p_Required As Boolean, _
    ByRef p_Logs() As String) As Boolean

    Dim fld As DAO.Field

    On Error GoTo Fail
    Set fld = p_Db.TableDefs(p_TableName).Fields(p_FieldName)
    If fld.Type <> p_Type Then GoTo Fail
    If p_Required And Not fld.Required Then GoTo Fail

    AssertField = True
    Exit Function

Fail:
    p_Logs(UBound(p_Logs)) = "Field contract failed: " & p_TableName & "." & p_FieldName
    AssertField = False
End Function

Private Function AssertIndexFields( _
    ByVal p_Db As DAO.Database, _
    ByVal p_TableName As String, _
    ByVal p_IndexName As String, _
    ByVal p_ExpectedFieldName As String, _
    ByVal p_Unique As Boolean, _
    ByRef p_Logs() As String) As Boolean

    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim fieldFound As Boolean
    Dim extraFieldCount As Long

    On Error GoTo Fail
    Set idx = p_Db.TableDefs(p_TableName).Indexes(p_IndexName)
    If idx.Unique <> p_Unique Then GoTo Fail

    extraFieldCount = 0
    fieldFound = False
    For Each fld In idx.Fields
        If StrComp(fld.Name, p_ExpectedFieldName, vbTextCompare) = 0 Then
            fieldFound = True
        Else
            extraFieldCount = extraFieldCount + 1
        End If
    Next fld

    If Not fieldFound Or extraFieldCount > 0 Then GoTo Fail

    AssertIndexFields = True
    Exit Function

Fail:
    p_Logs(UBound(p_Logs)) = "Index fields contract failed: " & p_TableName & "." & p_IndexName & _
                            " expected exactly '" & p_ExpectedFieldName & "'"
    AssertIndexFields = False
End Function

Private Function AssertIndex( _
    ByVal p_Db As DAO.Database, _
    ByVal p_TableName As String, _
    ByVal p_IndexName As String, _
    ByVal p_Unique As Boolean, _
    ByRef p_Logs() As String) As Boolean

    Dim idx As DAO.Index

    On Error GoTo Fail
    Set idx = p_Db.TableDefs(p_TableName).Indexes(p_IndexName)
    If idx.Unique <> p_Unique Then GoTo Fail

    AssertIndex = True
    Exit Function

Fail:
    p_Logs(UBound(p_Logs)) = "Index contract failed: " & p_TableName & "." & p_IndexName
    AssertIndex = False
End Function

Private Function AssertPrimaryKey( _
    ByVal p_Db As DAO.Database, _
    ByVal p_TableName As String, _
    ByVal p_FieldName As String, _
    ByRef p_Logs() As String) As Boolean

    Dim idx As DAO.Index
    Dim fld As DAO.Field

    On Error GoTo Fail
    For Each idx In p_Db.TableDefs(p_TableName).Indexes
        If idx.Primary Then
            If idx.Fields.Count = 1 Then
                For Each fld In idx.Fields
                    If StrComp(fld.Name, p_FieldName, vbTextCompare) = 0 Then
                        AssertPrimaryKey = True
                        Exit Function
                    End If
                Next fld
            End If
        End If
    Next idx

Fail:
    p_Logs(UBound(p_Logs)) = "Primary key contract failed: " & p_TableName & "." & p_FieldName
    AssertPrimaryKey = False
End Function

Private Function AssertNoRelationshipsForTable( _
    ByVal p_Db As DAO.Database, _
    ByVal p_TableName As String, _
    ByRef p_Logs() As String) As Boolean

    Dim rel As DAO.Relation

    On Error GoTo Fail
    For Each rel In p_Db.Relations
        If StrComp(rel.Table, p_TableName, vbTextCompare) = 0 Or StrComp(rel.ForeignTable, p_TableName, vbTextCompare) = 0 Then
            p_Logs(UBound(p_Logs)) = "Relationship contract failed: unexpected relation " & rel.Name & " touches " & p_TableName
            AssertNoRelationshipsForTable = False
            Exit Function
        End If
    Next rel

    AssertNoRelationshipsForTable = True
    Exit Function

Fail:
    p_Logs(UBound(p_Logs)) = "Relationship contract failed for: " & p_TableName
    AssertNoRelationshipsForTable = False
End Function

Private Function JsonOK(ByVal p_Ok As Boolean, ByVal p_Value As String, ByRef p_Logs() As String) As String
    If p_Ok Then
        JsonOK = BuildJsonOk(p_Value, p_Logs)
    Else
        JsonOK = BuildJsonFail(p_Value, p_Logs)
    End If
End Function
