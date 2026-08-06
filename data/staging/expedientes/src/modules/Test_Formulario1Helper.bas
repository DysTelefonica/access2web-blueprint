Attribute VB_Name = "Test_Formulario1Helper"
Option Compare Database
Option Explicit

' Test_Formulario1Helper — canonical JSON TDD atoms for Phase 3.4 / PR41.

Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

Private Function PayloadOf(ByVal p_Json As String) As Object
    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(p_Json)
    Set PayloadOf = parsed("payload")
End Function

Public Function Test_Formulario1Helper_BuildSetScript_EscapesApostrophe() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String
    Dim script As String
    script = modFormulario1Helper.Formulario1_BuildFirebaseSetScript("Bob's", errMsg)
    If InStr(script, "Bob\'s") = 0 Then Err.Raise 1001, , "expected escaped apostrophe"
    Test_Formulario1Helper_BuildSetScript_EscapesApostrophe = BuildOk("script-escaped", logs)
    Exit Function
EH:
    Test_Formulario1Helper_BuildSetScript_EscapesApostrophe = BuildFail(Err.Description, logs)
End Function

Public Function Test_Formulario1Helper_BuildTempFilePath_HappyTempDir() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String
    If modFormulario1Helper.Formulario1_BuildTempFilePath("C:\Temp", errMsg) <> "C:\Temp\firebaseRealtime.html" Then Err.Raise 1001, , "expected firebase html path"
    Test_Formulario1Helper_BuildTempFilePath_HappyTempDir = BuildOk("temp-path", logs)
    Exit Function
EH:
    Test_Formulario1Helper_BuildTempFilePath_HappyTempDir = BuildFail(Err.Description, logs)
End Function

Public Function Test_Formulario1Helper_BuildFirebaseHtml_IncludesRealtimeHook() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String
    Dim html As String
    html = modFormulario1Helper.Formulario1_BuildFirebaseHtml(errMsg)
    If InStr(html, "window.external.UpdateRealtime") = 0 Then Err.Raise 1001, , "expected external callback"
    Test_Formulario1Helper_BuildFirebaseHtml_IncludesRealtimeHook = BuildOk("firebase-html", logs)
    Exit Function
EH:
    Test_Formulario1Helper_BuildFirebaseHtml_IncludesRealtimeHook = BuildFail(Err.Description, logs)
End Function

Public Function Test_Formulario1Helper_RealtimePayload_ReturnsValue() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String
    Dim payload As Object
    Set payload = PayloadOf(modFormulario1Helper.Formulario1_RealtimePayload("hello", errMsg))
    If CStr(payload("value")) <> "hello" Then Err.Raise 1001, , "expected payload value"
    Test_Formulario1Helper_RealtimePayload_ReturnsValue = BuildOk("realtime-payload", logs)
    Exit Function
EH:
    Test_Formulario1Helper_RealtimePayload_ReturnsValue = BuildFail(Err.Description, logs)
End Function

Public Function Test_Formulario1Helper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    On Error GoTo EH
    logs(0) = Test_Formulario1Helper_BuildSetScript_EscapesApostrophe()
    logs(1) = Test_Formulario1Helper_BuildTempFilePath_HappyTempDir()
    logs(2) = Test_Formulario1Helper_BuildFirebaseHtml_IncludesRealtimeHook()
    logs(3) = Test_Formulario1Helper_RealtimePayload_ReturnsValue()
    Test_Formulario1Helper_RunAll = BuildOk("4 atoms executed", logs)
    Exit Function
EH:
    Test_Formulario1Helper_RunAll = BuildFail(Err.Description, logs)
End Function
