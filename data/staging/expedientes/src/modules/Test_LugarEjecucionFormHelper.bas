Attribute VB_Name = "Test_LugarEjecucionFormHelper"
Option Compare Database
Option Explicit

' Test_LugarEjecucionFormHelper — pure-data atoms for Phase 3.5 / PR42.

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

Private Function StubLugar(ByVal p_Lugar As String, ByVal p_Desc As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("LugarEjecucion") = p_Lugar
    d("DESCRIPCION") = p_Desc
    Set StubLugar = d
End Function

Public Function Test_LugarEjecucionFormHelper_FormOpenState_EditTitle() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modLugarEjecucionFormHelper.LugarEjecucionForm_FormOpenState(True, True, errMsg))
    If CBool(payload("registrarEnabled")) <> True Then Err.Raise 1001, , "expected enabled"
    If CStr(payload("title")) <> "EDICIÓN DE LUGAR EJECUCIÓN" Then Err.Raise 1002, , "expected edit title"
    Test_LugarEjecucionFormHelper_FormOpenState_EditTitle = BuildOk("state", logs)
    Exit Function
EH:
    Test_LugarEjecucionFormHelper_FormOpenState_EditTitle = BuildFail(Err.Description, logs)
End Function

Public Function Test_LugarEjecucionFormHelper_BuildValues_MapsFields() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, values As Object
    Set values = modLugarEjecucionFormHelper.LugarEjecucionForm_BuildValues("Madrid", "Desc", errMsg)
    If CStr(values("LugarEjecucion")) <> "Madrid" Then Err.Raise 1001, , "expected LugarEjecucion"
    If CStr(values("DESCRIPCION")) <> "Desc" Then Err.Raise 1002, , "expected desc"
    Test_LugarEjecucionFormHelper_BuildValues_MapsFields = BuildOk("values", logs)
    Exit Function
EH:
    Test_LugarEjecucionFormHelper_BuildValues_MapsFields = BuildFail(Err.Description, logs)
End Function

Public Function Test_LugarEjecucionFormHelper_RegisterDecision_NoChangesRejected() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, json As String
    json = modLugarEjecucionFormHelper.LugarEjecucionForm_RegisterDecision(False, False, errMsg)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected fail JSON"
    Test_LugarEjecucionFormHelper_RegisterDecision_NoChangesRejected = BuildOk("no-changes", logs)
    Exit Function
EH:
    Test_LugarEjecucionFormHelper_RegisterDecision_NoChangesRejected = BuildFail(Err.Description, logs)
End Function

Public Function Test_LugarEjecucionFormHelper_RegisterDecision_NewRaisesAlta() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modLugarEjecucionFormHelper.LugarEjecucionForm_RegisterDecision(True, True, errMsg))
    If CStr(payload("eventName")) <> "Alta" Then Err.Raise 1001, , "expected Alta"
    Test_LugarEjecucionFormHelper_RegisterDecision_NewRaisesAlta = BuildOk("alta", logs)
    Exit Function
EH:
    Test_LugarEjecucionFormHelper_RegisterDecision_NewRaisesAlta = BuildFail(Err.Description, logs)
End Function
