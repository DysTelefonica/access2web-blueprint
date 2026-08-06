Attribute VB_Name = "Test_E2EExportService"
Option Compare Database
Option Explicit

Public Function Test_E2EExport_AutomaticSelectsPendingOnly() As String
    Dim p_Error As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim cleanupNeeded As Boolean
    Dim pendingId As Long
    Dim exportedId As Long
    Dim candidateIdsCsv As String
    Dim actualCsv As String
    Dim logs(0 To 6) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_E2EExport_AutomaticSelectsPendingOnly = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    p_Error = ""
    pendingId = E2EExportFixtureBaseId() + 1
    exportedId = E2EExportFixtureBaseId() + 2
    candidateIdsCsv = CStr(pendingId) & "," & CStr(exportedId)

    logs(0) = "Arrange: setup sandbox and inspect real schema before seed"
    If Not SetupE2EBatchSchemaSandbox(errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_AutomaticSelectsPendingOnly = BuildJsonFail("sandbox setup failed: " & errLocal, logs)
        GoTo Cleanup
    End If
    cleanupNeeded = True

    If Not AssertE2EExportHashColumns(errLocal) Then
        p_Error = errLocal
        Test_E2EExport_AutomaticSelectsPendingOnly = BuildJsonFail("schema gate failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(1) = "Arrange: seed one pending and one unchanged exported expediente"
    If Not SeedE2EExportExpedienteFixture(pendingId, 9101, "HASH-PENDING-001", "", "EXP-E2E-AUTO-PENDING", errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_AutomaticSelectsPendingOnly = BuildJsonFail("pending seed failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    If Not SeedE2EExportExpedienteFixture(exportedId, 9102, "HASH-EXPORTED-001", "HASH-EXPORTED-001", "EXP-E2E-AUTO-EXPORTED", errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_AutomaticSelectsPendingOnly = BuildJsonFail("exported seed failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(2) = "Act: preview automatic export selection scoped to controlled fixture candidates"
    actualCsv = PreviewAutomaticE2EExportSelectionCsv(errLocal, candidateIdsCsv)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_AutomaticSelectsPendingOnly = BuildJsonFail("automatic preview failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(3) = "Assert: exact controlled candidate set contains only the pending fixture"
    If actualCsv = CStr(pendingId) And CsvValueCount(actualCsv) = 1 Then
        Test_E2EExport_AutomaticSelectsPendingOnly = BuildJsonOk("automatic export selects exactly one controlled pending fixture", logs)
    Else
        Test_E2EExport_AutomaticSelectsPendingOnly = BuildJsonFail("expected exactly " & CStr(pendingId) & " from candidates " & candidateIdsCsv & ", got '" & actualCsv & "'", logs)
    End If
    GoTo Cleanup

HandleError:
    p_Error = "Test_E2EExport_AutomaticSelectsPendingOnly: " & Err.Description
    logs(4) = p_Error
    Test_E2EExport_AutomaticSelectsPendingOnly = BuildJsonFail(p_Error, logs)

Cleanup:
    On Error Resume Next
    logs(5) = "Cleanup: delete deterministic export fixtures"
    If cleanupNeeded Then
        cleanupErr = ""
        If Not TeardownE2EExportFixture(cleanupErr) Or cleanupErr <> "" Then
            p_Error = cleanupErr
            Test_E2EExport_AutomaticSelectsPendingOnly = BuildJsonFail("cleanup failed: " & cleanupErr, logs)
        End If
    End If
    logs(6) = "Cleanup complete"
End Function

Public Function Test_E2EExport_ManualUsesRightSideSelectionOnly() As String
    Dim p_Error As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim cleanupNeeded As Boolean
    Dim selectedId As Long
    Dim leftOnlyId As Long
    Dim actualCsv As String
    Dim logs(0 To 6) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_E2EExport_ManualUsesRightSideSelectionOnly = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    p_Error = ""
    selectedId = E2EExportFixtureBaseId() + 11
    leftOnlyId = E2EExportFixtureBaseId() + 12

    logs(0) = "Arrange: setup sandbox and inspect real schema before seed"
    If Not SetupE2EBatchSchemaSandbox(errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_ManualUsesRightSideSelectionOnly = BuildJsonFail("sandbox setup failed: " & errLocal, logs)
        GoTo Cleanup
    End If
    cleanupNeeded = True

    If Not AssertE2EExportHashColumns(errLocal) Then
        p_Error = errLocal
        Test_E2EExport_ManualUsesRightSideSelectionOnly = BuildJsonFail("schema gate failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(1) = "Arrange: seed selected and available-left expediente rows"
    If Not SeedE2EExportExpedienteFixture(selectedId, 9111, "HASH-MANUAL-SELECTED", "", "EXP-E2E-MANUAL-SELECTED", errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_ManualUsesRightSideSelectionOnly = BuildJsonFail("selected seed failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    If Not SeedE2EExportExpedienteFixture(leftOnlyId, 9112, "HASH-MANUAL-LEFT", "", "EXP-E2E-MANUAL-LEFT", errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_ManualUsesRightSideSelectionOnly = BuildJsonFail("left-only seed failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    If Not SetupE2ESelectionFixture("qa.export.manual", "S-EXPORT-MANUAL", CStr(selectedId), errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_ManualUsesRightSideSelectionOnly = BuildJsonFail("manual selection seed failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(2) = "Act: preview manual export selection for user/session"
    actualCsv = PreviewManualE2EExportSelectionCsv("qa.export.manual", "S-EXPORT-MANUAL", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_ManualUsesRightSideSelectionOnly = BuildJsonFail("manual preview failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(3) = "Assert: only right-side selected expediente is exported"
    If actualCsv = CStr(selectedId) Then
        Test_E2EExport_ManualUsesRightSideSelectionOnly = BuildJsonOk("manual export uses right-side selection only", logs)
    Else
        Test_E2EExport_ManualUsesRightSideSelectionOnly = BuildJsonFail("expected only " & CStr(selectedId) & ", got '" & actualCsv & "'", logs)
    End If
    GoTo Cleanup

HandleError:
    p_Error = "Test_E2EExport_ManualUsesRightSideSelectionOnly: " & Err.Description
    logs(4) = p_Error
    Test_E2EExport_ManualUsesRightSideSelectionOnly = BuildJsonFail(p_Error, logs)

Cleanup:
    On Error Resume Next
    logs(5) = "Cleanup: delete deterministic export fixtures"
    If cleanupNeeded Then
        cleanupErr = ""
        If Not TeardownE2EExportFixture(cleanupErr) Or cleanupErr <> "" Then
            p_Error = cleanupErr
            Test_E2EExport_ManualUsesRightSideSelectionOnly = BuildJsonFail("cleanup failed: " & cleanupErr, logs)
        End If
    End If
    logs(6) = "Cleanup complete"
End Function

Public Function Test_E2EExport_JsonMetaIncludesExportId() As String
    Dim p_Error As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim cleanupNeeded As Boolean
    Dim batchId As Long
    Dim jsonText As String
    Dim hasExportId As Boolean
    Dim logs(0 To 5) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_E2EExport_JsonMetaIncludesExportId = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    p_Error = ""

    logs(0) = "Arrange: setup sandbox batch schema"
    If Not SetupE2EBatchSchemaSandbox(errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_JsonMetaIncludesExportId = BuildJsonFail("sandbox setup failed: " & errLocal, logs)
        GoTo Cleanup
    End If
    cleanupNeeded = True

    logs(1) = "Arrange: create deterministic batch header"
    If Not CreateE2EExportBatchFixture("qa.export.meta", "S-EXPORT-META", batchId, errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_JsonMetaIncludesExportId = BuildJsonFail("batch seed failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(2) = "Act: generate batch JSON"
    jsonText = GenerateE2EExportJsonForBatch(batchId, CStr(E2EExportFixtureBaseId() + 21), errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_JsonMetaIncludesExportId = BuildJsonFail("batch JSON generation failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(3) = "Assert: JSON meta contains IDExportacion or IDBatch matching persisted batch"
    hasExportId = (InStr(1, jsonText, Chr$(34) & "IDExportacion" & Chr$(34), vbTextCompare) > 0 Or _
                   InStr(1, jsonText, Chr$(34) & "IDBatch" & Chr$(34), vbTextCompare) > 0) And _
                  (InStr(1, jsonText, CStr(batchId), vbTextCompare) > 0)

    If hasExportId Then
        Test_E2EExport_JsonMetaIncludesExportId = BuildJsonOk("JSON meta includes export/batch identifier", logs)
    Else
        Test_E2EExport_JsonMetaIncludesExportId = BuildJsonFail("JSON meta missing export/batch identifier for IDBatch=" & CStr(batchId), logs)
    End If
    GoTo Cleanup

HandleError:
    p_Error = "Test_E2EExport_JsonMetaIncludesExportId: " & Err.Description
    logs(4) = p_Error
    Test_E2EExport_JsonMetaIncludesExportId = BuildJsonFail(p_Error, logs)

Cleanup:
    On Error Resume Next
    logs(5) = "Cleanup: delete deterministic batch fixtures"
    If cleanupNeeded Then
        cleanupErr = ""
        If Not TeardownE2EExportFixture(cleanupErr) Or cleanupErr <> "" Then
            p_Error = cleanupErr
            Test_E2EExport_JsonMetaIncludesExportId = BuildJsonFail("cleanup failed: " & cleanupErr, logs)
        End If
    End If
End Function

Public Function Test_E2EExport_SuccessRecordsHistoryAndAdvancesHash() As String
    Dim p_Error As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim cleanupNeeded As Boolean
    Dim pendingId As Long
    Dim resultJson As String
    Dim parsed As Object
    Dim selectedIds As String
    Dim detailRows As Long
    Dim exportedHash As String
    Dim outputPath As String
    Dim logs(0 To 6) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_E2EExport_SuccessRecordsHistoryAndAdvancesHash = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    p_Error = ""
    pendingId = E2EExportFixtureBaseId() + 31

    logs(0) = "Arrange: setup sandbox and deterministic pending expediente"
    If Not SetupE2EBatchSchemaSandbox(errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_SuccessRecordsHistoryAndAdvancesHash = BuildJsonFail("sandbox setup failed: " & errLocal, logs)
        GoTo Cleanup
    End If
    cleanupNeeded = True

    If Not AssertE2EExportHashColumns(errLocal) Then
        p_Error = errLocal
        Test_E2EExport_SuccessRecordsHistoryAndAdvancesHash = BuildJsonFail("schema gate failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(1) = "Arrange: seed pending expediente with current hash HASH-SUCCESS-001"
    If Not SeedE2EExportExpedienteFixture(pendingId, 9131, "HASH-SUCCESS-001", "", "EXP-E2E-SUCCESS-PENDING", errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_SuccessRecordsHistoryAndAdvancesHash = BuildJsonFail("pending seed failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(2) = "Act: run export service success path scoped to the controlled fixture candidate"
    resultJson = ExportE2EAutomatic("qa.export.success", "S-EXPORT-SUCCESS", BuildE2EExportTempPath("success"), "ok", errLocal, CStr(pendingId))
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_SuccessRecordsHistoryAndAdvancesHash = BuildJsonFail("automatic export failed: " & errLocal, logs)
        GoTo Cleanup
    End If
    Set parsed = JsonConverter.ParseJson(resultJson)
    selectedIds = CStr(parsed("selectedIds"))
    outputPath = CStr(parsed("outputPath"))

    detailRows = CountE2EExportDetailRows(pendingId, "Exitoso", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_SuccessRecordsHistoryAndAdvancesHash = BuildJsonFail("detail count failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    exportedHash = ReadE2EExportHashUltima(pendingId, errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_SuccessRecordsHistoryAndAdvancesHash = BuildJsonFail("hash read failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(3) = "Assert: file exists, only controlled fixture selected, success detail exists, and hash advances"
    If CBool(parsed("ok")) = True And selectedIds = CStr(pendingId) And E2EExportFileExists(outputPath) And detailRows = 1 And exportedHash = "HASH-SUCCESS-001" Then
        Test_E2EExport_SuccessRecordsHistoryAndAdvancesHash = BuildJsonOk("success writes file, records history and advances hash", logs)
    Else
        Test_E2EExport_SuccessRecordsHistoryAndAdvancesHash = BuildJsonFail("expected selectedIds " & CStr(pendingId) & ", output file, one success detail, and exported hash HASH-SUCCESS-001", logs)
    End If
    GoTo Cleanup

HandleError:
    p_Error = "Test_E2EExport_SuccessRecordsHistoryAndAdvancesHash: " & Err.Description
    logs(4) = p_Error
    Test_E2EExport_SuccessRecordsHistoryAndAdvancesHash = BuildJsonFail(p_Error, logs)

Cleanup:
    On Error Resume Next
    logs(5) = "Cleanup: delete deterministic export fixtures"
    If cleanupNeeded Then
        cleanupErr = ""
        If Not TeardownE2EExportFixture(cleanupErr) Or cleanupErr <> "" Then
            p_Error = cleanupErr
            Test_E2EExport_SuccessRecordsHistoryAndAdvancesHash = BuildJsonFail("cleanup failed: " & cleanupErr, logs)
        End If
    End If
    DeleteE2EExportTempPath "success"
    logs(6) = "Cleanup complete"
End Function

Public Function Test_E2EExport_FileFailureDoesNotMarkExported() As String
    Dim p_Error As String
    Dim errLocal As String
    Dim cleanupErr As String
    Dim cleanupNeeded As Boolean
    Dim failureId As Long
    Dim resultJson As String
    Dim parsed As Object
    Dim successDetailRows As Long
    Dim exportedHash As String
    Dim logs(0 To 6) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_E2EExport_FileFailureDoesNotMarkExported = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    p_Error = ""
    failureId = E2EExportFixtureBaseId() + 41

    logs(0) = "Arrange: setup sandbox and deterministic failure-path expediente"
    If Not SetupE2EBatchSchemaSandbox(errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_FileFailureDoesNotMarkExported = BuildJsonFail("sandbox setup failed: " & errLocal, logs)
        GoTo Cleanup
    End If
    cleanupNeeded = True

    If Not AssertE2EExportHashColumns(errLocal) Then
        p_Error = errLocal
        Test_E2EExport_FileFailureDoesNotMarkExported = BuildJsonFail("schema gate failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(1) = "Arrange: seed pending expediente with old exported hash"
    If Not SeedE2EExportExpedienteFixture(failureId, 9141, "HASH-FAILURE-001", "HASH-FAILURE-000", "EXP-E2E-FAILURE-PENDING", errLocal) Or errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_FileFailureDoesNotMarkExported = BuildJsonFail("failure seed failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(2) = "Act: run export service with forced file writer failure"
    resultJson = ExportE2EAutomatic("qa.export.failure", "S-EXPORT-FAILURE", BuildE2EExportTempPath("failure"), "force-fail", errLocal, CStr(failureId))
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_FileFailureDoesNotMarkExported = BuildJsonFail("automatic export failed: " & errLocal, logs)
        GoTo Cleanup
    End If
    Set parsed = JsonConverter.ParseJson(resultJson)

    successDetailRows = CountE2EExportDetailRows(failureId, "Exitoso", errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_FileFailureDoesNotMarkExported = BuildJsonFail("detail count failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    exportedHash = ReadE2EExportHashUltima(failureId, errLocal)
    If errLocal <> "" Then
        p_Error = errLocal
        Test_E2EExport_FileFailureDoesNotMarkExported = BuildJsonFail("hash read failed: " & errLocal, logs)
        GoTo Cleanup
    End If

    logs(3) = "Assert: failure result has no success detail and leaves exported hash unchanged"
    If CBool(parsed("ok")) = False And successDetailRows = 0 And exportedHash = "HASH-FAILURE-000" Then
        Test_E2EExport_FileFailureDoesNotMarkExported = BuildJsonOk("file failure does not mark exported", logs)
    Else
        Test_E2EExport_FileFailureDoesNotMarkExported = BuildJsonFail("failure path marked exported or failed to report failure", logs)
    End If
    GoTo Cleanup

HandleError:
    p_Error = "Test_E2EExport_FileFailureDoesNotMarkExported: " & Err.Description
    logs(4) = p_Error
    Test_E2EExport_FileFailureDoesNotMarkExported = BuildJsonFail(p_Error, logs)

Cleanup:
    On Error Resume Next
    logs(5) = "Cleanup: delete deterministic export fixtures"
    If cleanupNeeded Then
        cleanupErr = ""
        If Not TeardownE2EExportFixture(cleanupErr) Or cleanupErr <> "" Then
            p_Error = cleanupErr
            Test_E2EExport_FileFailureDoesNotMarkExported = BuildJsonFail("cleanup failed: " & cleanupErr, logs)
        End If
    End If
    DeleteE2EExportTempPath "failure"
    logs(6) = "Cleanup complete"
End Function

' ------------------------------------------------------------------------------------------
' e2e export writer keeps literal UTF-8 unicode chars in the on-disk file
' (closes #79 — regression guard for the ANSI \uXXXX writer bug).
' Compares BYTES, not VBA strings, so the assert is robust to whatever
' encoding Access applies to the .bas source at import time.
' ------------------------------------------------------------------------------------------
Public Function Test_E2EExportService_WriteJsonUtf8File_WritesUtf8LiteralChars() As String
    Dim logs(0 To 3) As String
    Dim tmpPath As String
    Dim jsonEscaped As String
    Dim fileNum As Integer
    Dim fileLen As Long
    Dim fileBytes() As Byte
    Dim patternSi(0 To 2) As Byte
    Dim patternAnio(0 To 3) As Byte
    Dim patternEscape() As Byte
    Dim hasSiUtf8 As Boolean
    Dim hasAnioUtf8 As Boolean
    Dim hasUnicodeEscape As Boolean
    Dim writeErr As String
    Dim writeOk As Boolean

    Test_E2EExportService_WriteJsonUtf8File_WritesUtf8LiteralChars = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    tmpPath = Environ$("TEMP") & "\expedientes_e2e_utf8_write_json_utf8_file.json"
    jsonEscaped = "{""texto"":""S\u00ed"",""a\u00f1o"":2026}"

    logs(0) = "Arrange: temp path " & tmpPath
    logs(1) = "Act: call WriteJsonUtf8File with JSON containing \u00ed and \u00f1 escapes"

    writeErr = ""
    writeOk = WriteJsonUtf8File(tmpPath, jsonEscaped, writeErr)
    If Not writeOk Then
        Test_E2EExportService_WriteJsonUtf8File_WritesUtf8LiteralChars = BuildJsonFail("WriteJsonUtf8File failed: " & writeErr, logs)
        GoTo Cleanup
    End If

    logs(2) = "Assert: read file bytes via VBA Open For Binary + Get"

    fileNum = FreeFile
    Open tmpPath For Binary Access Read As #fileNum
    fileLen = LOF(fileNum)
    If fileLen > 0 Then
        ReDim fileBytes(0 To fileLen - 1) As Byte
        Get #fileNum, , fileBytes
    Else
        ReDim fileBytes(0 To 0) As Byte
    End If
    Close #fileNum

    patternSi(0) = &H53: patternSi(1) = &HC3: patternSi(2) = &HAD
    patternAnio(0) = &H61: patternAnio(1) = &HC3: patternAnio(2) = &HB1: patternAnio(3) = &H6F
    patternEscape = StrConv("\u00", vbFromUnicode)

    hasSiUtf8 = ContainsByteSeq(fileBytes, patternSi)
    hasAnioUtf8 = ContainsByteSeq(fileBytes, patternAnio)
    hasUnicodeEscape = ContainsByteSeq(fileBytes, patternEscape)

    logs(3) = "Assert: bytes contain UTF-8 Sí (53 C3 AD), UTF-8 año (61 C3 B1 6F), and no literal '\u00' ASCII bytes"

    If Not hasSiUtf8 Then
        Test_E2EExportService_WriteJsonUtf8File_WritesUtf8LiteralChars = BuildJsonFail("expected UTF-8 bytes 53 C3 AD for 'Sí' in file", logs)
    ElseIf Not hasAnioUtf8 Then
        Test_E2EExportService_WriteJsonUtf8File_WritesUtf8LiteralChars = BuildJsonFail("expected UTF-8 bytes 61 C3 B1 6F for 'año' in file", logs)
    ElseIf hasUnicodeEscape Then
        Test_E2EExportService_WriteJsonUtf8File_WritesUtf8LiteralChars = BuildJsonFail("literal '\u00' ASCII bytes found in file (writer did not unescape)", logs)
    Else
        Test_E2EExportService_WriteJsonUtf8File_WritesUtf8LiteralChars = BuildJsonOk("utf-8 file keeps literal unicode bytes, no \uXXXX escapes", logs)
    End If
    GoTo Cleanup

HandleError:
    Test_E2EExportService_WriteJsonUtf8File_WritesUtf8LiteralChars = BuildJsonFail("utf-8 file assertion failed: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    If Len(tmpPath) > 0 Then Kill tmpPath
End Function

' ------------------------------------------------------------------------------------------
' Returns True if haystack contains pattern as a contiguous byte sequence.
' Robust to any source-file encoding: compares bytes, not VBA strings.
' ------------------------------------------------------------------------------------------
Private Function ContainsByteSeq(ByRef haystack() As Byte, ByRef pattern() As Byte) As Boolean
    Dim hLen As Long, pLen As Long, i As Long, j As Long, match As Boolean
    ContainsByteSeq = False
    hLen = UBound(haystack) - LBound(haystack) + 1
    pLen = UBound(pattern) - LBound(pattern) + 1
    If pLen <= 0 Or pLen > hLen Then Exit Function
    For i = LBound(haystack) To UBound(haystack) - pLen + 1
        match = True
        For j = LBound(pattern) To UBound(pattern)
            If haystack(i + j) <> pattern(j) Then
                match = False
                Exit For
            End If
        Next j
        If match Then
            ContainsByteSeq = True
            Exit Function
        End If
    Next i
End Function

Private Function AssertE2EExportHashColumns(ByRef p_Error As String) As Boolean
    Dim db As DAO.Database
    p_Error = ""

    If Not EnsureE2EExportTraceabilitySchema(p_Error) Then Exit Function

    Set db = getdb(p_Error)
    If p_Error <> "" Then Exit Function
    If db Is Nothing Then
        p_Error = "AssertE2EExportHashColumns: getdb returned Nothing"
        Exit Function
    End If

    If Not TableFieldExists(db, "TbExpedientes", "HashActual") Then
        p_Error = "TbExpedientes.HashActual missing in real sandbox schema"
        Exit Function
    End If

    If Not TableFieldExists(db, "TbExpedientes", "HashUltimaExportacion") Then
        p_Error = "TbExpedientes.HashUltimaExportacion missing in real sandbox schema"
        Exit Function
    End If

    AssertE2EExportHashColumns = True
End Function

Private Function CountE2EExportDetailRows(ByVal p_IDExpediente As Long, ByVal p_Estado As String, ByRef p_Error As String) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    On Error GoTo EH
    p_Error = ""

    Set db = getdb(p_Error)
    If p_Error <> "" Then Exit Function
    If db Is Nothing Then
        p_Error = "CountE2EExportDetailRows: getdb returned Nothing"
        Exit Function
    End If

    Set rs = db.OpenRecordset("SELECT COUNT(*) AS Cnt FROM TbE2EExportBatchDetalle WHERE IDExpediente=" & _
                              CLng(p_IDExpediente) & " AND Estado='" & SqlStr(p_Estado) & "';", dbOpenSnapshot)
    If Not rs.EOF Then CountE2EExportDetailRows = CLng(Nz(rs.Fields("Cnt").value, 0))
    rs.Close
    Set rs = Nothing
    Exit Function

EH:
    p_Error = "CountE2EExportDetailRows: " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

Private Function ReadE2EExportHashUltima(ByVal p_IDExpediente As Long, ByRef p_Error As String) As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    On Error GoTo EH
    p_Error = ""

    Set db = getdb(p_Error)
    If p_Error <> "" Then Exit Function
    If db Is Nothing Then
        p_Error = "ReadE2EExportHashUltima: getdb returned Nothing"
        Exit Function
    End If

    Set rs = db.OpenRecordset("SELECT HashUltimaExportacion FROM TbExpedientes WHERE IDExpediente=" & CLng(p_IDExpediente) & ";", dbOpenSnapshot)
    If Not rs.EOF Then ReadE2EExportHashUltima = Nz(rs.Fields("HashUltimaExportacion").value, "")
    rs.Close
    Set rs = Nothing
    Exit Function

EH:
    p_Error = "ReadE2EExportHashUltima: " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

Private Function BuildE2EExportTempPath(ByVal p_Token As String) As String
    Dim fsoLocal As Object
    Dim basePath As String

    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    basePath = Environ$("TEMP")
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    BuildE2EExportTempPath = basePath & "expedientes_e2e_export_" & p_Token
    If Not fsoLocal.FolderExists(BuildE2EExportTempPath) Then fsoLocal.CreateFolder BuildE2EExportTempPath
End Function

Private Sub DeleteE2EExportTempPath(ByVal p_Token As String)
    Dim fsoLocal As Object
    Dim folderPath As String
    Dim basePath As String

    On Error Resume Next
    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    basePath = Environ$("TEMP")
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    folderPath = basePath & "expedientes_e2e_export_" & p_Token
    If fsoLocal.FolderExists(folderPath) Then fsoLocal.DeleteFolder folderPath, True
End Sub

Private Function E2EExportFileExists(ByVal p_FilePath As String) As Boolean
    Dim fsoLocal As Object

    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    E2EExportFileExists = fsoLocal.FileExists(p_FilePath)
End Function

Private Function CsvValueCount(ByVal p_Csv As String) As Long
    If Trim$(p_Csv) = "" Then Exit Function
    CsvValueCount = UBound(Split(p_Csv, ",")) + 1
End Function

Private Function SqlStr(ByVal p_Value As String) As String
    SqlStr = Replace(Nz(p_Value, ""), "'", "''")
End Function

Private Function TableFieldExists(ByVal p_Db As DAO.Database, ByVal p_TableName As String, ByVal p_FieldName As String) As Boolean
    Dim fld As DAO.Field

    On Error Resume Next
    Set fld = p_Db.TableDefs(p_TableName).Fields(p_FieldName)
    TableFieldExists = (Err.Number = 0 And Not fld Is Nothing)
    Err.Clear
    Set fld = Nothing
End Function



