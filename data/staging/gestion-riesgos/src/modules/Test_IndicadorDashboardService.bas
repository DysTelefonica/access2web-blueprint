Attribute VB_Name = "Test_IndicadorDashboardService"
Option Compare Database
Option Explicit

Public Function Test_IndicadorDashboard_BuildOpenArgs_CanonicalContract() As String
    Dim logs(0 To 3) As String
    Dim openArgs As String
    Dim errorText As String

    On Error GoTo EH
    logs(0) = "1. Arrange: year 2025 and published semester labels"
    logs(1) = "2. Act: IndicadorDashboard_BuildOpenArgs"
    logs(2) = "3. Assert: canonical ANIO=<year>;SEM=<value> contract"

    openArgs = IndicadorDashboard_BuildOpenArgs("2025", "S1", errorText)
    If openArgs <> "ANIO=2025;SEM=1" Or Len(errorText) > 0 Then GoTo Failed
    openArgs = IndicadorDashboard_BuildOpenArgs("2025", "S2", errorText)
    If openArgs <> "ANIO=2025;SEM=2" Or Len(errorText) > 0 Then GoTo Failed
    openArgs = IndicadorDashboard_BuildOpenArgs("2025", "Anual", errorText)
    If openArgs <> "ANIO=2025;SEM=" Or Len(errorText) > 0 Then GoTo Failed

    logs(3) = "4. Assert PASS: producer emits the canonical contract"
    Test_IndicadorDashboard_BuildOpenArgs_CanonicalContract = TestCore_BuildOk("canonical_openargs", logs)
    Exit Function
Failed:
    Test_IndicadorDashboard_BuildOpenArgs_CanonicalContract = TestCore_BuildFail("Unexpected OpenArgs: " & openArgs & " / " & errorText, logs)
    Exit Function
EH:
    Test_IndicadorDashboard_BuildOpenArgs_CanonicalContract = TestCore_BuildFail(Err.Description, logs)
End Function

Public Function Test_IndicadorDashboard_BuildOpenArgs_RejectsInvalidInputs() As String
    Dim logs(0 To 3) As String
    Dim openArgs As String
    Dim errorText As String

    On Error GoTo EH
    logs(0) = "1. Arrange: invalid year and semester inputs"
    logs(1) = "2. Act: IndicadorDashboard_BuildOpenArgs"
    logs(2) = "3. Assert: empty result and explicit message"

    openArgs = IndicadorDashboard_BuildOpenArgs("abc", "S1", errorText)
    If Len(openArgs) > 0 Or Len(errorText) = 0 Then GoTo Failed
    errorText = ""
    openArgs = IndicadorDashboard_BuildOpenArgs("2025", "Q3", errorText)
    If Len(openArgs) > 0 Or Len(errorText) = 0 Then GoTo Failed

    logs(3) = "4. Assert PASS: invalid UI input never opens the child form"
    Test_IndicadorDashboard_BuildOpenArgs_RejectsInvalidInputs = TestCore_BuildOk("invalid_rejected", logs)
    Exit Function
Failed:
    Test_IndicadorDashboard_BuildOpenArgs_RejectsInvalidInputs = TestCore_BuildFail("Invalid input was not rejected", logs)
    Exit Function
EH:
    Test_IndicadorDashboard_BuildOpenArgs_RejectsInvalidInputs = TestCore_BuildFail(Err.Description, logs)
End Function

Public Function Test_IndicadorDashboard_BuildExportSql_UsesSuccessfulContext() As String
    Dim logs(0 To 3) As String
    Dim sql As String
    Dim errorText As String

    On Error GoTo EH
    logs(0) = "1. Arrange: last successful calculation context"
    logs(1) = "2. Act: IndicadorDashboard_BuildExportSql for tile 6"
    logs(2) = "3. Assert: SQL uses cached ids and dates and requires success"

    sql = IndicadorDashboard_BuildExportSql(itVigentesEnPeriodo, #1/1/2025#, #12/31/2025#, "900700", True, errorText)
    If Len(sql) = 0 Or InStr(1, sql, "900700", vbBinaryCompare) = 0 Then GoTo FailedFirstCall
    errorText = ""
    sql = IndicadorDashboard_BuildExportSql(itVigentesEnPeriodo, #1/1/2025#, #12/31/2025#, "900700", False, errorText)
    If Len(sql) > 0 Or Len(errorText) = 0 Then GoTo FailedSecondCall

    logs(3) = "4. Assert PASS: export is bound to successful calculation context"
    Test_IndicadorDashboard_BuildExportSql_UsesSuccessfulContext = TestCore_BuildOk("cached_export_context", logs)
    Exit Function
FailedFirstCall:
    Test_IndicadorDashboard_BuildExportSql_UsesSuccessfulContext = TestCore_BuildFail( _
        "ExportFirstCall: errorText='" & errorText & "', sqlLen=" & Len(sql) & ", sql[0..120]='" & Left(sql, 120) & "'", _
        logs)
    Exit Function
FailedSecondCall:
    Test_IndicadorDashboard_BuildExportSql_UsesSuccessfulContext = TestCore_BuildFail( _
        "ExportSecondCall: errorText='" & errorText & "', sql='" & sql & "'", _
        logs)
    Exit Function
EH:
    Test_IndicadorDashboard_BuildExportSql_UsesSuccessfulContext = TestCore_BuildFail(Err.Description, logs)
End Function

Public Function Test_IndicadorDashboard_CanExport_MapsSixTiles() As String
    Dim logs(0 To 3) As String
    Dim values() As Long

    On Error GoTo EH
    ReDim values(1 To 6)
    logs(0) = "1. Arrange: only tile 6 has a positive value"
    logs(1) = "2. Act: IndicadorDashboard_CanExport"
    logs(2) = "3. Assert: tile 6 enabled and tile 1 disabled"
    values(6) = 2

    If Not IndicadorDashboard_CanExport(values, itVigentesEnPeriodo) Then GoTo Failed
    If IndicadorDashboard_CanExport(values, itIdentificados) Then GoTo Failed

    logs(3) = "4. Assert PASS: all six tile indexes are supported"
    Test_IndicadorDashboard_CanExport_MapsSixTiles = TestCore_BuildOk("six_tiles", logs)
    Exit Function
Failed:
    Test_IndicadorDashboard_CanExport_MapsSixTiles = TestCore_BuildFail("Six-tile export mapping failed", logs)
    Exit Function
EH:
    Test_IndicadorDashboard_CanExport_MapsSixTiles = TestCore_BuildFail(Err.Description, logs)
End Function
