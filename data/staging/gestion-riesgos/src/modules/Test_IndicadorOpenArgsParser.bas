Attribute VB_Name = "Test_IndicadorOpenArgsParser"
Option Compare Database
Option Explicit

Private Function AssertParse( _
    ByVal p_OpenArgs As String, _
    ByVal p_ExpectedOk As Boolean, _
    ByVal p_ExpectedYear As Long, _
    ByVal p_ExpectedSemester As String, _
    ByVal p_ErrorContains As String _
) As String
    Dim actualYear As Long
    Dim actualSemester As String
    Dim actualError As String
    Dim actualOk As Boolean

    actualOk = ParseIndicadorOpenArgs(p_OpenArgs, actualYear, actualSemester, actualError)

    If actualOk <> p_ExpectedOk Then
        AssertParse = "Unexpected result for '" & p_OpenArgs & "'"
        Exit Function
    End If
    If actualYear <> p_ExpectedYear Then
        AssertParse = "Unexpected year for '" & p_OpenArgs & "': " & CStr(actualYear)
        Exit Function
    End If
    If actualSemester <> p_ExpectedSemester Then
        AssertParse = "Unexpected semester for '" & p_OpenArgs & "': '" & actualSemester & "'"
        Exit Function
    End If
    If Len(p_ErrorContains) = 0 Then
        If Len(actualError) > 0 Then AssertParse = "Unexpected error for '" & p_OpenArgs & "': " & actualError
    ElseIf InStr(1, actualError, p_ErrorContains, vbTextCompare) = 0 Then
        AssertParse = "Expected error containing '" & p_ErrorContains & "', got '" & actualError & "'"
    End If
End Function

Public Function Test_IndicadorOpenArgsParser_CanonicalContracts() As String
    Dim logs(0 To 3) As String
    Dim failure As String

    On Error GoTo EH
    logs(0) = "1. Arrange: canonical named OpenArgs for S1, S2 and annual"
    logs(1) = "2. Act: ParseIndicadorOpenArgs for each published contract"
    logs(2) = "3. Assert: year and canonical semester values"

    failure = AssertParse("ANIO=2025;SEM=1", True, 2025, "1", "")
    If Len(failure) = 0 Then failure = AssertParse("ANIO=2025;SEM=2", True, 2025, "2", "")
    If Len(failure) = 0 Then failure = AssertParse("ANIO=2025;SEM=", True, 2025, "", "")
    If Len(failure) > 0 Then
        Test_IndicadorOpenArgsParser_CanonicalContracts = TestCore_BuildFail(failure, logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: all canonical variants parsed"
    Test_IndicadorOpenArgsParser_CanonicalContracts = TestCore_BuildOk("canonical_contracts", logs)
    Exit Function
EH:
    Test_IndicadorOpenArgsParser_CanonicalContracts = TestCore_BuildFail(Err.Description, logs)
End Function

Public Function Test_IndicadorOpenArgsParser_MalformedInputFailsExplicitly() As String
    Dim logs(0 To 3) As String
    Dim failure As String

    On Error GoTo EH
    logs(0) = "1. Arrange: empty, non-numeric year and invalid semester inputs"
    logs(1) = "2. Act: ParseIndicadorOpenArgs"
    logs(2) = "3. Assert: explicit failure with reset outputs and useful error"

    failure = AssertParse("", False, 0, "", "OpenArgs")
    If Len(failure) = 0 Then failure = AssertParse("ANIO=abc;SEM=1", False, 0, "", "ANIO")
    If Len(failure) = 0 Then failure = AssertParse("ANIO=2025;SEM=3", False, 0, "", "Semestre")
    If Len(failure) > 0 Then
        Test_IndicadorOpenArgsParser_MalformedInputFailsExplicitly = TestCore_BuildFail(failure, logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: malformed inputs fail explicitly"
    Test_IndicadorOpenArgsParser_MalformedInputFailsExplicitly = TestCore_BuildOk("malformed_rejected", logs)
    Exit Function
EH:
    Test_IndicadorOpenArgsParser_MalformedInputFailsExplicitly = TestCore_BuildFail(Err.Description, logs)
End Function

Public Function Test_IndicadorOpenArgsParser_LegacyPipeCompatibility() As String
    Dim logs(0 To 3) As String
    Dim failure As String

    On Error GoTo EH
    logs(0) = "1. Arrange: published defensive pipe inputs"
    logs(1) = "2. Act: ParseIndicadorOpenArgs"
    logs(2) = "3. Assert: legacy S1, S2 and annual remain compatible"

    failure = AssertParse("2025|S1", True, 2025, "1", "")
    If Len(failure) = 0 Then failure = AssertParse("2025|2", True, 2025, "2", "")
    If Len(failure) = 0 Then failure = AssertParse("2025|", True, 2025, "", "")
    If Len(failure) > 0 Then
        Test_IndicadorOpenArgsParser_LegacyPipeCompatibility = TestCore_BuildFail(failure, logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: legacy pipe variants remain compatible"
    Test_IndicadorOpenArgsParser_LegacyPipeCompatibility = TestCore_BuildOk("legacy_compatible", logs)
    Exit Function
EH:
    Test_IndicadorOpenArgsParser_LegacyPipeCompatibility = TestCore_BuildFail(Err.Description, logs)
End Function

Public Function Test_IndicadorOpenArgsParser_FailureResetsOutputs() As String
    Dim logs(0 To 3) As String
    Dim actualYear As Long
    Dim actualSemester As String
    Dim actualError As String
    Dim actualOk As Boolean

    On Error GoTo EH
    logs(0) = "1. Arrange: stale output values and malformed named input"
    logs(1) = "2. Act: ParseIndicadorOpenArgs"
    logs(2) = "3. Assert: no stale year or semester survives failure"
    actualYear = 2030
    actualSemester = "2"

    actualOk = ParseIndicadorOpenArgs("ANIO=;SEM=1", actualYear, actualSemester, actualError)
    If actualOk Or actualYear <> 0 Or actualSemester <> "" Or Len(actualError) = 0 Then
        Test_IndicadorOpenArgsParser_FailureResetsOutputs = TestCore_BuildFail("Failure must reset every output", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: failure outputs are deterministic"
    Test_IndicadorOpenArgsParser_FailureResetsOutputs = TestCore_BuildOk("outputs_reset", logs)
    Exit Function
EH:
    Test_IndicadorOpenArgsParser_FailureResetsOutputs = TestCore_BuildFail(Err.Description, logs)
End Function
