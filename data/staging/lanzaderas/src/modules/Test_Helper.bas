Attribute VB_Name = "Test_Helper"
Option Compare Database
Option Explicit

Private Function TestJsonValue(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        TestJsonValue = "null"
        Exit Function
    End If

    Select Case VarType(value)
        Case vbBoolean
            TestJsonValue = LCase$(CStr(value))
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            TestJsonValue = Replace(CStr(value), ",", ".")
        Case Else
            TestJsonValue = """" & TestEscapeJsonString(CStr(value)) & """"
    End Select
End Function

Private Function TestJsonStringArray(ByRef logs() As String) As String
    On Error GoTo EmptyLogs

    Dim i As Long
    Dim parts As String
    For i = LBound(logs) To UBound(logs)
        If Len(logs(i)) > 0 Then
            If Len(parts) > 0 Then parts = parts & ","
            parts = parts & """" & TestEscapeJsonString(logs(i)) & """"
        End If
    Next i

    TestJsonStringArray = "[" & parts & "]"
    Exit Function

EmptyLogs:
    TestJsonStringArray = "[]"
End Function

Public Function BuildJsonOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildJsonOk = "{""ok"":true,""value"":" & TestJsonValue(value) & ",""payload"":null,""error"":null,""logs"":" & TestJsonStringArray(logs) & "}"
End Function

Public Function BuildJsonFail(ByVal errorMsg As String, ByRef logs() As String) As String
    BuildJsonFail = "{""ok"":false,""value"":null,""payload"":null,""error"":""" & TestEscapeJsonString(errorMsg) & """,""logs"":" & TestJsonStringArray(logs) & "}"
End Function

Public Function TestEscapeJsonString(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    TestEscapeJsonString = s
End Function

Public Function ReadProjectFile(ByVal relativePath As String, Optional ByRef p_Error As String = "") As String
    On Error GoTo EH
    p_Error = ""

    Dim absolutePath As String
    absolutePath = CurrentProject.Path & "\" & relativePath

    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(absolutePath) Then
        p_Error = "Archivo no encontrado: " & absolutePath
        Exit Function
    End If

    Set ts = fso.OpenTextFile(absolutePath, 1, False)
    ReadProjectFile = ts.ReadAll
    ts.Close
    Exit Function

EH:
    p_Error = "ReadProjectFile: " & Err.Number & " - " & Err.Description
End Function

Public Function ExtractProcedures(ByVal manifestText As String) As Collection
    Dim rx As Object
    Dim matches As Object
    Dim m As Object
    Dim procedures As New Collection

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = """procedure""\s*:\s*""([^""]+)"""

    Set matches = rx.Execute(manifestText)
    For Each m In matches
        procedures.Add CStr(m.SubMatches(0))
    Next m

    Set ExtractProcedures = procedures
End Function

Public Function ExtractJsonPropertyValues(ByVal jsonText As String, ByVal propertyName As String) As Collection
    Dim rx As Object
    Dim matches As Object
    Dim m As Object
    Dim values As New Collection

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = Chr$(34) & propertyName & Chr$(34) & "\s*:\s*" & Chr$(34) & "([^" & Chr$(34) & "]+)" & Chr$(34)

    Set matches = rx.Execute(jsonText)
    For Each m In matches
        values.Add CStr(m.SubMatches(0))
    Next m

    Set ExtractJsonPropertyValues = values
End Function

Public Function IsRunAllProcedure(ByVal procedureName As String) As Boolean
    IsRunAllProcedure = (Right$(procedureName, 7) = "_RunAll")
End Function

Public Function CollectionContains(ByVal values As Collection, ByVal target As String) As Boolean
    Dim i As Long
    For i = 1 To values.Count
        If CStr(values(i)) = target Then
            CollectionContains = True
            Exit Function
        End If
    Next i
End Function

Public Function JoinCollection(ByVal values As Collection) As String
    Dim i As Long
    Dim result As String
    For i = 1 To values.Count
        If result <> "" Then result = result & ", "
        result = result & CStr(values(i))
    Next i
    JoinCollection = result
End Function

Public Function JsonResultOk(ByVal jsonResult As String) As Boolean
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = "^\s*\{.*?""ok""\s*:\s*true"

    JsonResultOk = rx.Test(jsonResult)
End Function

Public Function ExtractJsonError(ByVal jsonResult As String) As String
    Dim values As Collection
    Set values = ExtractJsonPropertyValues(jsonResult, "error")

    If values.Count > 0 Then
        If LCase$(Trim$(CStr(values(1)))) <> "null" Then
            ExtractJsonError = CStr(values(1))
        End If
    End If
End Function

Public Function BuildChildFailureSummary(ByVal procedureName As String, ByVal jsonResult As String) As String
    Dim childError As String
    childError = ExtractJsonError(jsonResult)

    If childError = "" Then
        BuildChildFailureSummary = procedureName
    Else
        BuildChildFailureSummary = procedureName & ": " & childError
    End If
End Function

Private Function ResolveBackendSandbox(ByRef p_BackendPath As String, _
                                       ByRef p_BackendPassword As String, _
                                       Optional ByRef p_Error As String = "") As Boolean
    On Error GoTo EH
    p_Error = ""
    p_BackendPath = ""
    p_BackendPassword = ""

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT TOP 1 BackendSandbox, PasswordBackend FROM TbConfiguracionBackends", dbOpenSnapshot)

    If rs.EOF Then
        p_Error = "TESTS BLOCKED: TbConfiguracionBackends no tiene configuración"
        GoTo Cleanup
    End If

    p_BackendPath = Trim$(Nz(rs.Fields("BackendSandbox").value, ""))
    p_BackendPassword = Nz(rs.Fields("PasswordBackend").value, "")
    If p_BackendPassword = "" Then
        p_BackendPassword = Environ$("ACCESS_VBA_PASSWORD")
    End If

    If p_BackendPath = "" Then
        p_Error = "TESTS BLOCKED: BackendSandbox está vacío"
        GoTo Cleanup
    End If

    ResolveBackendSandbox = True

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
EH:
    p_Error = "ResolveBackendSandbox: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function

Private Function ValidateBackendSandbox(ByVal p_BackendPath As String, _
                                        ByVal p_BackendPassword As String, _
                                        Optional ByRef p_Error As String = "") As Boolean
    On Error GoTo EH
    p_Error = ""

    Dim localFso As Object
    Set localFso = CreateObject("Scripting.FileSystemObject")
    If Not localFso.FileExists(p_BackendPath) Then
        p_Error = "TESTS BLOCKED: BackendSandbox no encontrado: " & p_BackendPath
        Exit Function
    End If

    Dim sandboxDb As DAO.Database
    Set sandboxDb = DBEngine.Workspaces(0).OpenDatabase(p_BackendPath, False, False, ";PWD=" & p_BackendPassword)
    sandboxDb.Close
    Set sandboxDb = Nothing

    ValidateBackendSandbox = True
    Exit Function
EH:
    p_Error = "TESTS BLOCKED: BackendSandbox no alcanzable: " & Err.Number & " - " & Err.Description
End Function

Public Function ForceLocalBackend(Optional ByRef p_Error As String = "") As Boolean
    On Error GoTo EH
    p_Error = ""

    Dim backendPath As String
    Dim backendPassword As String
    If Not ResolveBackendSandbox(backendPath, backendPassword, p_Error) Then Exit Function
    If Not ValidateBackendSandbox(backendPath, backendPassword, p_Error) Then Exit Function

    SetTestingMode True, backendPath, p_Error
    If p_Error <> "" Then Exit Function

    ResetGlobals p_Error
    If p_Error <> "" Then Exit Function

    LeeConfiguracionLocal p_Error
    If p_Error <> "" Then Exit Function

    ForceLocalBackend = True
    Exit Function
EH:
    p_Error = "ForceLocalBackend: " & Err.Number & " - " & Err.Description
End Function

Public Function ResetTestSession(Optional ByRef p_Error As String = "") As Boolean
    On Error GoTo EH
    p_Error = ""

    SetTestingMode False, "", p_Error
    If p_Error <> "" Then Exit Function

    ResetGlobals p_Error
    If p_Error <> "" Then Exit Function

    Application.TempVars.RemoveAll
    ResetTestSession = True
    Exit Function
EH:
    p_Error = "ResetTestSession: " & Err.Number & " - " & Err.Description
End Function

Public Function SafeTempVarValue(ByVal tempVarName As String) As String
    On Error GoTo MissingTempVar
    SafeTempVarValue = CStr(Nz(Application.TempVars(tempVarName).value, ""))
    Exit Function

MissingTempVar:
    SafeTempVarValue = ""
End Function



