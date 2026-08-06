Attribute VB_Name = "Test_HarnessContract"
Option Compare Database
Option Explicit

Private Const MANIFEST_RELATIVE_PATH As String = "tests\tests.vba.json"
Private Const TEST_MODULES_RELATIVE_PATH As String = "src\modules"

Public Function Test_HarnessManifest_ProceduresAreUnqualifiedUnique() As String
    Dim logs(0 To 3) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_HarnessManifest_ProceduresAreUnqualifiedUnique = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    Dim manifestText As String
    Dim procedures As Collection
    Dim errorText As String

    logs(0) = "Arrange: read Dysflow VBA manifest"
    manifestText = ReadUtf8Text(ProjectFilePath(MANIFEST_RELATIVE_PATH))

    logs(1) = "Act: extract manifest procedure names"
    Set procedures = ExtractManifestProcedures(manifestText)

    logs(2) = "Assert: procedure names are unqualified, unique, and non-generic"
    If Not AssertManifestProcedureNames(procedures, errorText) Then
        Test_HarnessManifest_ProceduresAreUnqualifiedUnique = BuildJsonFail(errorText, logs)
        Exit Function
    End If

    Test_HarnessManifest_ProceduresAreUnqualifiedUnique = BuildJsonOk(CStr(procedures.Count) & " manifest procedures comply", logs)
    Exit Function

HandleError:
    logs(3) = "Error: " & Err.Description
    Test_HarnessManifest_ProceduresAreUnqualifiedUnique = BuildJsonFail("Test_HarnessManifest_ProceduresAreUnqualifiedUnique: " & Err.Description, logs)
End Function

Public Function Test_HarnessManifest_ProceduresAreZeroArgJsonFunctions() As String
    Dim logs(0 To 4) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_HarnessManifest_ProceduresAreZeroArgJsonFunctions = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    Dim manifestText As String
    Dim sourceText As String
    Dim procedures As Collection
    Dim errorText As String

    logs(0) = "Arrange: read manifest and exported VBA test sources"
    manifestText = ReadUtf8Text(ProjectFilePath(MANIFEST_RELATIVE_PATH))
    sourceText = ReadAllTestModuleSource(ProjectFilePath(TEST_MODULES_RELATIVE_PATH))

    logs(1) = "Act: extract manifest procedures"
    Set procedures = ExtractManifestProcedures(manifestText)

    logs(2) = "Assert: each manifested procedure is a zero-argument Public Function returning String"
    If Not AssertManifestProceduresAreZeroArgFunctions(procedures, sourceText, errorText) Then
        Test_HarnessManifest_ProceduresAreZeroArgJsonFunctions = BuildJsonFail(errorText, logs)
        Exit Function
    End If

    logs(3) = "Assert: canonical JSON helpers expose ok/value/payload/error/logs contract"
    Test_HarnessManifest_ProceduresAreZeroArgJsonFunctions = BuildJsonOk(CStr(procedures.Count) & " procedures expose the Dysflow contract", logs)
    Exit Function

HandleError:
    logs(4) = "Error: " & Err.Description
    Test_HarnessManifest_ProceduresAreZeroArgJsonFunctions = BuildJsonFail("Test_HarnessManifest_ProceduresAreZeroArgJsonFunctions: " & Err.Description, logs)
End Function

Public Function Test_HarnessTestModules_DoNotExposeLegacyOutputOrRealPaths() As String
    Dim logs(0 To 4) As String
    ' Default: fail. Solo se override si los asserts pasan.
    Test_HarnessTestModules_DoNotExposeLegacyOutputOrRealPaths = BuildJsonFail("test did not complete", logs)

    On Error GoTo HandleError

    Dim errorText As String
    Dim modulesPath As String

    modulesPath = ProjectFilePath(TEST_MODULES_RELATIVE_PATH)

    logs(0) = "Arrange: enumerate exported Test*.bas modules, including legacy non-Test_ modules"
    logs(1) = "Act: scan runnable test modules for legacy output and environment coupling"

    If Not AssertTestModulesAvoidLegacyPatterns(modulesPath, errorText) Then
        Test_HarnessTestModules_DoNotExposeLegacyOutputOrRealPaths = BuildJsonFail(errorText, logs)
        Exit Function
    End If

    logs(2) = "Assert: no Sub-style public test entrypoints, debug output, UI prompts, real IDs, or user paths remain"
    Test_HarnessTestModules_DoNotExposeLegacyOutputOrRealPaths = BuildJsonOk("all Test*.bas runnable modules comply", logs)
    Exit Function

HandleError:
    logs(4) = "Error: " & Err.Description
    Test_HarnessTestModules_DoNotExposeLegacyOutputOrRealPaths = BuildJsonFail("Test_HarnessTestModules_DoNotExposeLegacyOutputOrRealPaths: " & Err.Description, logs)
End Function

Private Function ProjectFilePath(ByVal p_RelativePath As String) As String
    ProjectFilePath = Application.CurrentProject.Path & "\" & p_RelativePath
End Function

Private Function ReadUtf8Text(ByVal p_Path As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile p_Path
        ReadUtf8Text = .ReadText(-1)
        .Close
    End With
    Set stream = Nothing
End Function

Private Function ReadAllTestModuleSource(ByVal p_ModulesPath As String) As String
    Dim fsoLocal As Object
    Dim folder As Object
    Dim fileItem As Object
    Dim sourceText As String

    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    Set folder = fsoLocal.GetFolder(p_ModulesPath)

    For Each fileItem In folder.Files
        If IsRunnableTestModule(CStr(fileItem.Name)) Then
                sourceText = sourceText & vbCrLf & ReadUtf8Text(CStr(fileItem.Path))
        End If
    Next fileItem

    ReadAllTestModuleSource = sourceText
End Function

Private Function AssertTestModulesAvoidLegacyPatterns(ByVal p_ModulesPath As String, ByRef p_Error As String) As Boolean
    Dim fsoLocal As Object
    Dim folder As Object
    Dim fileItem As Object
    Dim sourceText As String
    Dim sourceNoComments As String

    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    Set folder = fsoLocal.GetFolder(p_ModulesPath)
    p_Error = ""

    For Each fileItem In folder.Files
        If IsRunnableTestModule(CStr(fileItem.Name)) Then
            sourceText = ReadUtf8Text(CStr(fileItem.Path))
            sourceNoComments = StripVbaCommentLines(sourceText)

            If InStr(1, sourceNoComments, "Public " & "Sub ", vbTextCompare) > 0 Then
                p_Error = CStr(fileItem.Name) & " exposes a Sub-style public test entrypoint; tests must be zero-argument Public Function As String"
                Exit Function
            End If
            If InStr(1, sourceNoComments, "Debug" & ".Print", vbTextCompare) > 0 Then
                p_Error = CStr(fileItem.Name) & " uses debug-print output; test output must be JSON logs"
                Exit Function
            End If
            If InStr(1, sourceNoComments, "Msg" & "Box", vbTextCompare) > 0 Or InStr(1, sourceNoComments, "Input" & "Box", vbTextCompare) > 0 Then
                p_Error = CStr(fileItem.Name) & " uses blocking UI output"
                Exit Function
            End If
            If InStr(1, sourceNoComments, "C:\" & "Users\", vbTextCompare) > 0 Then
                p_Error = CStr(fileItem.Name) & " hardcodes a user profile path"
                Exit Function
            End If
            If InStr(1, sourceNoComments, "IDExpediente IN (" & "1072, 402)", vbTextCompare) > 0 Then
                p_Error = CStr(fileItem.Name) & " hardcodes real expediente IDs"
                Exit Function
            End If
        End If
    Next fileItem

    AssertTestModulesAvoidLegacyPatterns = True
End Function

Private Function IsRunnableTestModule(ByVal p_FileName As String) As Boolean
    Dim lowerName As String

    lowerName = LCase$(p_FileName)
    If Right$(lowerName, 4) <> ".bas" Then Exit Function
    If Left$(lowerName, 4) <> "test" Then Exit Function

    Select Case lowerName
        Case "test.bas", "testhelper.bas", "testfixtures.bas"
            IsRunnableTestModule = False
        Case Else
            IsRunnableTestModule = True
    End Select
End Function

Private Function StripVbaCommentLines(ByVal p_SourceText As String) As String
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim result As String

    lines = Split(Replace(p_SourceText, vbCrLf, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        lineText = LTrim$(lines(i))
        If Left$(lineText, 1) <> "'" Then
            result = result & lineText & vbLf
        End If
    Next i

    StripVbaCommentLines = result
End Function

Private Function ExtractManifestProcedures(ByVal p_ManifestText As String) As Collection
    Dim procedures As Collection
    Dim searchFrom As Long
    Dim keyPos As Long
    Dim valueStart As Long
    Dim valueEnd As Long

    Set procedures = New Collection
    searchFrom = 1

    Do
        keyPos = InStr(searchFrom, p_ManifestText, """procedure""", vbTextCompare)
        If keyPos = 0 Then Exit Do

        valueStart = InStr(keyPos, p_ManifestText, ":", vbBinaryCompare)
        valueStart = InStr(valueStart, p_ManifestText, """", vbBinaryCompare) + 1
        valueEnd = InStr(valueStart, p_ManifestText, """", vbBinaryCompare)

        If valueStart > 1 And valueEnd > valueStart Then
            procedures.Add Mid$(p_ManifestText, valueStart, valueEnd - valueStart)
        End If

        searchFrom = valueEnd + 1
    Loop

    Set ExtractManifestProcedures = procedures
End Function

Private Function AssertManifestProcedureNames(ByVal p_Procedures As Collection, ByRef p_Error As String) As Boolean
    Dim seen As Object
    Dim procedureName As Variant

    Set seen = CreateObject("Scripting.Dictionary")
    p_Error = ""

    For Each procedureName In p_Procedures
        If InStr(1, CStr(procedureName), ".", vbBinaryCompare) > 0 Then
            p_Error = "Manifest procedure must be unqualified: " & CStr(procedureName)
            Exit Function
        End If
        If InStr(1, CStr(procedureName), "(", vbBinaryCompare) > 0 Or InStr(1, CStr(procedureName), ")", vbBinaryCompare) > 0 Then
            p_Error = "Manifest procedure must not include parentheses: " & CStr(procedureName)
            Exit Function
        End If
        If StrComp(CStr(procedureName), "RunAll", vbTextCompare) = 0 Then
            p_Error = "Manifest procedure must not use generic RunAll"
            Exit Function
        End If
        If seen.Exists(LCase$(CStr(procedureName))) Then
            p_Error = "Manifest procedure duplicated: " & CStr(procedureName)
            Exit Function
        End If
        seen.Add LCase$(CStr(procedureName)), True
    Next procedureName

    AssertManifestProcedureNames = True
End Function

Private Function AssertManifestProceduresAreZeroArgFunctions( _
    ByVal p_Procedures As Collection, _
    ByVal p_SourceText As String, _
    ByRef p_Error As String) As Boolean

    Dim procedureName As Variant
    Dim expectedSignature As String

    p_Error = ""

    For Each procedureName In p_Procedures
        expectedSignature = "Public Function " & CStr(procedureName) & "() As String"
        If InStr(1, p_SourceText, expectedSignature, vbTextCompare) = 0 Then
            p_Error = "Manifest procedure is not a zero-argument String function in exported Test_*.bas sources: " & CStr(procedureName)
            Exit Function
        End If
    Next procedureName

    AssertManifestProceduresAreZeroArgFunctions = True
End Function
