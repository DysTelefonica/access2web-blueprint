Attribute VB_Name = "Test_TareasHelper"
Option Compare Database
Option Explicit

' Test_TareasHelper — canonical JSON TDD atoms for Phase 3.4 / PR41.

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

Public Function Test_TareasHelper_ListHeader_Happy() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modTareasHelper.Tareas_ListHeader(errMsg))
    If InStr(CStr(payload("rowSource")), "Nemotécnico") = 0 Then Err.Raise 1001, , "expected tasks header"
    Test_TareasHelper_ListHeader_Happy = BuildOk("header", logs)
    Exit Function
EH:
    Test_TareasHelper_ListHeader_Happy = BuildFail(Err.Description, logs)
End Function

Public Function Test_TareasHelper_TaskType_SadNonNumeric() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, json As String
    json = modTareasHelper.Tareas_TaskTypeIsValid("x", errMsg)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false"
    Test_TareasHelper_TaskType_SadNonNumeric = BuildOk("non-numeric-rejected", logs)
    Exit Function
EH:
    Test_TareasHelper_TaskType_SadNonNumeric = BuildFail(Err.Description, logs)
End Function

Public Function Test_TareasHelper_ReportAction_SadUnknown() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, json As String
    json = modTareasHelper.Tareas_ReportAction(9999, errMsg)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected unknown task failure"
    Test_TareasHelper_ReportAction_SadUnknown = BuildOk("unknown-rejected", logs)
    Exit Function
EH:
    Test_TareasHelper_ReportAction_SadUnknown = BuildFail(Err.Description, logs)
End Function

Public Function Test_TareasHelper_SelectionState_EnablesActions() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modTareasHelper.Tareas_SelectionState("EXP-1", "", errMsg))
    If CBool(payload("enableDetail")) <> True Then Err.Raise 1001, , "expected detail enabled"
    If CBool(payload("enableInforme")) <> True Then Err.Raise 1002, , "expected report enabled"
    Test_TareasHelper_SelectionState_EnablesActions = BuildOk("selection", logs)
    Exit Function
EH:
    Test_TareasHelper_SelectionState_EnablesActions = BuildFail(Err.Description, logs)
End Function
