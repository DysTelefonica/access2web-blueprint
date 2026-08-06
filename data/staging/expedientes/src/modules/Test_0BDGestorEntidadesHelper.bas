Attribute VB_Name = "Test_0BDGestorEntidadesHelper"
Option Compare Database
Option Explicit

' Test_0BDGestorEntidadesHelper — pure-data atoms for Phase 3.5 / PR42.

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

Private Function StubOptions() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("LUG") = "Lugares;Ejecución"
    d("GRA") = "Grados"
    Set StubOptions = d
End Function

Private Function StubFormMap() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("LUG") = "FormLugarEjecucionGestion"
    Set StubFormMap = d
End Function

Public Function Test_0BDGestorEntidadesHelper_ListRowSource_SanitizesSemicolon() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(mod0BDGestorEntidadesHelper.GestorEntidades_ListRowSource(StubOptions(), errMsg))
    If CLng(payload("count")) <> 2 Then Err.Raise 1001, , "expected count=2"
    If InStr(CStr(payload("rowSource")), "Lugares:Ejecución") = 0 Then Err.Raise 1002, , "expected semicolon sanitized"
    Test_0BDGestorEntidadesHelper_ListRowSource_SanitizesSemicolon = BuildOk("list", logs)
    Exit Function
EH:
    Test_0BDGestorEntidadesHelper_ListRowSource_SanitizesSemicolon = BuildFail(Err.Description, logs)
End Function

Public Function Test_0BDGestorEntidadesHelper_SelectionState_EmptyDisablesOpen() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(mod0BDGestorEntidadesHelper.GestorEntidades_SelectionState("", errMsg))
    If CBool(payload("enableOpen")) <> False Then Err.Raise 1001, , "expected disabled"
    Test_0BDGestorEntidadesHelper_SelectionState_EmptyDisablesOpen = BuildOk("empty", logs)
    Exit Function
EH:
    Test_0BDGestorEntidadesHelper_SelectionState_EmptyDisablesOpen = BuildFail(Err.Description, logs)
End Function

Public Function Test_0BDGestorEntidadesHelper_OpenDecision_MapsForm() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(mod0BDGestorEntidadesHelper.GestorEntidades_OpenDecision("LUG", StubFormMap(), errMsg))
    If CStr(payload("formName")) <> "FormLugarEjecucionGestion" Then Err.Raise 1001, , "expected mapped form"
    Test_0BDGestorEntidadesHelper_OpenDecision_MapsForm = BuildOk("mapped", logs)
    Exit Function
EH:
    Test_0BDGestorEntidadesHelper_OpenDecision_MapsForm = BuildFail(Err.Description, logs)
End Function

Public Function Test_0BDGestorEntidadesHelper_OpenDecision_RejectsUnknown() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, json As String
    json = mod0BDGestorEntidadesHelper.GestorEntidades_OpenDecision("XXX", StubFormMap(), errMsg)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected fail JSON"
    Test_0BDGestorEntidadesHelper_OpenDecision_RejectsUnknown = BuildOk("unknown", logs)
    Exit Function
EH:
    Test_0BDGestorEntidadesHelper_OpenDecision_RejectsUnknown = BuildFail(Err.Description, logs)
End Function
