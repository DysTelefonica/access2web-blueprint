Attribute VB_Name = "Test_RunnerProbe"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Test_RunnerProbe.bas
' RESPONSABILIDAD: Probes mínimos para aislar resolución Application.Run/Dysflow.
'
' Estas funciones no tocan datos ni comportamiento productivo. Solo devuelven JSON
' compacto para diagnosticar si el runner puede localizar procedimientos públicos
' y si un procedimiento público puede llamar internamente a GetPasswordDB().
' ==========================================================================

Public Function Test_RunnerProbe_Ping() As String
    Test_RunnerProbe_Ping = RunnerProbe_JsonOk("pong")
End Function

Public Function Test_RunnerProbe_ProjectName() As String
    On Error GoTo EH

    Test_RunnerProbe_ProjectName = "{" & RunnerProbe_Q() & "ok" & RunnerProbe_Q() & ":true," & _
        RunnerProbe_Q() & "value" & RunnerProbe_Q() & ":" & RunnerProbe_Q() & "project" & RunnerProbe_Q() & "," & _
        RunnerProbe_Q() & "payload" & RunnerProbe_Q() & ":{" & _
        RunnerProbe_JsonPair("currentProjectName", CurrentProject.Name) & "," & _
        RunnerProbe_JsonPair("currentProjectFullName", CurrentProject.FullName) & "," & _
        RunnerProbe_JsonPair("applicationName", Application.Name) & "}}"
    Exit Function

EH:
    Test_RunnerProbe_ProjectName = RunnerProbe_JsonError("Test_RunnerProbe_ProjectName", Err.Number, Err.Description)
End Function

Public Function Test_RunnerProbe_GetPasswordDB_Wrapper() As String
    On Error GoTo EH

    Dim passwordValue As String
    passwordValue = GetPasswordDB()

    Test_RunnerProbe_GetPasswordDB_Wrapper = "{" & RunnerProbe_Q() & "ok" & RunnerProbe_Q() & ":true," & _
        RunnerProbe_Q() & "value" & RunnerProbe_Q() & ":" & RunnerProbe_Q() & "GetPasswordDB" & RunnerProbe_Q() & "," & _
        RunnerProbe_Q() & "payload" & RunnerProbe_Q() & ":{" & _
        RunnerProbe_Q() & "length" & RunnerProbe_Q() & ":" & CStr(Len(passwordValue)) & "}}"
    Exit Function

EH:
    Test_RunnerProbe_GetPasswordDB_Wrapper = RunnerProbe_JsonError("GetPasswordDB", Err.Number, Err.Description)
End Function

Public Function Test_HarnessV24_BeginTestSession_Contract() As String
    On Error GoTo EH

    Dim logs() As String
    Dim errMsg As String
    logs = TestHelper.NewLogsArray(5)

    logs(0) = "1. Arrange: call canonical BeginTestSession with logs and error refs"
    If Not TestHelper.BeginTestSession(logs, errMsg) Then
        logs(1) = "2. RED: BeginTestSession did not complete canonical setup: " & errMsg
        Test_HarnessV24_BeginTestSession_Contract = TestHelper.BuildJsonFail(errMsg, logs)
        GoTo CleanExit
    End If

    logs(1) = "2. Assert: BeginTestSession returned Boolean True"
    If Not m_TestingMode Then
        logs(2) = "3. RED: m_TestingMode was not enabled"
        Test_HarnessV24_BeginTestSession_Contract = TestHelper.BuildJsonFail("m_TestingMode not enabled", logs)
        GoTo CleanExit
    End If

    If Len(Trim$(m_BackendSandboxURL)) = 0 Then
        logs(2) = "3. RED: sandbox URL was not cached"
        Test_HarnessV24_BeginTestSession_Contract = TestHelper.BuildJsonFail("sandbox URL not cached", logs)
        GoTo CleanExit
    End If

    logs(2) = "3. Assert: testing mode and sandbox URL are initialized"
    Test_HarnessV24_BeginTestSession_Contract = TestHelper.BuildJsonOk("begin_contract_ok", logs)

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_HarnessV24_BeginTestSession_Contract = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_HarnessV24_AssertSandboxBackend_Contract() As String
    On Error GoTo EH

    Dim logs() As String
    Dim errMsg As String
    Dim sandboxPath As String
    logs = TestHelper.NewLogsArray(5)

    logs(0) = "1. Arrange: call read-only AssertSandboxBackend"
    If Not TestHelper.AssertSandboxBackend(logs, sandboxPath, errMsg) Then
        logs(1) = "2. RED: sandbox assertion failed before any fixture write: " & errMsg
        Test_HarnessV24_AssertSandboxBackend_Contract = TestHelper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    If Len(Trim$(sandboxPath)) = 0 Then
        logs(1) = "2. RED: assertion passed without returning sandbox path"
        Test_HarnessV24_AssertSandboxBackend_Contract = TestHelper.BuildJsonFail("sandbox path not returned", logs)
        Exit Function
    End If

    logs(1) = "2. Assert: sandbox path returned"
    Test_HarnessV24_AssertSandboxBackend_Contract = TestHelper.BuildJsonOk("assert_sandbox_contract_ok", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_HarnessV24_AssertSandboxBackend_Contract = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_HarnessV24_EndTestSession_Contract() As String
    On Error GoTo EH

    Dim logs() As String
    Dim errMsg As String
    logs = TestHelper.NewLogsArray(5)

    logs(0) = "1. Arrange: set test routing state before teardown"
    m_TestingMode = True
    m_BackendSandboxURL = "C:\\temp\\condor_datos.accdb"
    m_PasswordBackend = "test"

    If Not TestHelper.EndTestSession(logs, errMsg) Then
        logs(1) = "2. RED: EndTestSession returned False: " & errMsg
        Test_HarnessV24_EndTestSession_Contract = TestHelper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    If m_TestingMode Then
        logs(1) = "2. RED: m_TestingMode remains True after EndTestSession"
        Test_HarnessV24_EndTestSession_Contract = TestHelper.BuildJsonFail("m_TestingMode not cleared", logs)
        Exit Function
    End If

    If Len(m_BackendSandboxURL) > 0 Then
        logs(1) = "2. RED: sandbox URL remains after EndTestSession"
        Test_HarnessV24_EndTestSession_Contract = TestHelper.BuildJsonFail("sandbox URL not cleared", logs)
        Exit Function
    End If

    logs(1) = "2. Assert: EndTestSession clears routing state"
    Test_HarnessV24_EndTestSession_Contract = TestHelper.BuildJsonOk("end_contract_ok", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_HarnessV24_EndTestSession_Contract = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Public Function Test_HarnessV24_ResetTestSession_Contract() As String
    On Error GoTo EH

    Dim logs() As String
    Dim errMsg As String
    logs = TestHelper.NewLogsArray(5)

    logs(0) = "1. Arrange: dirty test globals and TempVars"
    m_TestingMode = True
    m_BackendSandboxURL = "C:\\temp\\condor_datos.accdb"
    m_PasswordBackend = "test"
    Application.TempVars("BackendPathSandbox") = "dirty"
    Application.TempVars("BackendPathConfigurado") = "dirty"
    Application.TempVars("DatosEnLocal") = "dirty"

    Call TestHelper.ResetTestSession(errMsg)
    If Len(errMsg) > 0 Then
        logs(1) = "2. RED: ResetTestSession reported cleanup warning: " & errMsg
        Test_HarnessV24_ResetTestSession_Contract = TestHelper.BuildJsonFail(errMsg, logs)
        Exit Function
    End If

    If m_TestingMode Then
        logs(1) = "2. RED: m_TestingMode remains True"
        Test_HarnessV24_ResetTestSession_Contract = TestHelper.BuildJsonFail("m_TestingMode not cleared", logs)
        Exit Function
    End If

    If Len(m_BackendSandboxURL) > 0 Then
        logs(1) = "2. RED: m_BackendSandboxURL remains set"
        Test_HarnessV24_ResetTestSession_Contract = TestHelper.BuildJsonFail("m_BackendSandboxURL not cleared", logs)
        Exit Function
    End If

    logs(1) = "2. Assert: ResetTestSession is idempotent and clears test state"
    Test_HarnessV24_ResetTestSession_Contract = TestHelper.BuildJsonOk("reset_contract_ok", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_HarnessV24_ResetTestSession_Contract = TestHelper.BuildJsonFail(Err.Description, logs)
End Function

Private Function RunnerProbe_JsonOk(ByVal value As String) As String
    RunnerProbe_JsonOk = "{" & RunnerProbe_Q() & "ok" & RunnerProbe_Q() & ":true," & _
        RunnerProbe_Q() & "value" & RunnerProbe_Q() & ":" & RunnerProbe_Q() & RunnerProbe_EscapeJson(value) & RunnerProbe_Q() & "}"
End Function

Private Function RunnerProbe_JsonError(ByVal sourceName As String, ByVal errorNumber As Long, ByVal errorDescription As String) As String
    RunnerProbe_JsonError = "{" & RunnerProbe_Q() & "ok" & RunnerProbe_Q() & ":false," & _
        RunnerProbe_Q() & "value" & RunnerProbe_Q() & ":" & RunnerProbe_Q() & RunnerProbe_EscapeJson(sourceName) & RunnerProbe_Q() & "," & _
        RunnerProbe_Q() & "error" & RunnerProbe_Q() & ":{" & _
        RunnerProbe_Q() & "number" & RunnerProbe_Q() & ":" & CStr(errorNumber) & "," & _
        RunnerProbe_JsonPair("description", errorDescription) & "}}"
End Function

Private Function RunnerProbe_JsonPair(ByVal key As String, ByVal value As String) As String
    RunnerProbe_JsonPair = RunnerProbe_Q() & RunnerProbe_EscapeJson(key) & RunnerProbe_Q() & ":" & _
        RunnerProbe_Q() & RunnerProbe_EscapeJson(value) & RunnerProbe_Q()
End Function

Private Function RunnerProbe_EscapeJson(ByVal value As String) As String
    Dim escaped As String
    escaped = value
    escaped = Replace(escaped, "\", "\\")
    escaped = Replace(escaped, RunnerProbe_Q(), "\" & RunnerProbe_Q())
    escaped = Replace(escaped, vbCrLf, "\n")
    escaped = Replace(escaped, vbCr, "\n")
    escaped = Replace(escaped, vbLf, "\n")
    escaped = Replace(escaped, vbTab, "\t")
    RunnerProbe_EscapeJson = escaped
End Function

Private Function RunnerProbe_Q() As String
    RunnerProbe_Q = Chr$(34)
End Function
