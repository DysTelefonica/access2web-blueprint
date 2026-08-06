Attribute VB_Name = "Test_ExpedientesParaCambioTipoHelper"
Option Compare Database
Option Explicit

' Test_ExpedientesParaCambioTipoHelper — canonical JSON TDD atoms for Phase 3.4 / PR41.

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

Public Function Test_ExpedientesParaCambioTipoHelper_ListHeader_Happy() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modExpedientesParaCambioTipoHelper.ExpedientesParaCambioTipo_ListHeader(errMsg))
    If InStr(CStr(payload("rowSource")), "IDExpediente;Tipo") = 0 Then Err.Raise 1001, , "expected list header"
    Test_ExpedientesParaCambioTipoHelper_ListHeader_Happy = BuildOk("header", logs)
    Exit Function
EH:
    Test_ExpedientesParaCambioTipoHelper_ListHeader_Happy = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedientesParaCambioTipoHelper_SearchTerm_SadBlank() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, json As String
    json = modExpedientesParaCambioTipoHelper.ExpedientesParaCambioTipo_ValidateSearchTerm(" ", errMsg)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false"
    Test_ExpedientesParaCambioTipoHelper_SearchTerm_SadBlank = BuildOk("blank-rejected", logs)
    Exit Function
EH:
    Test_ExpedientesParaCambioTipoHelper_SearchTerm_SadBlank = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedientesParaCambioTipoHelper_SelectionState_EdgeEmpty() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modExpedientesParaCambioTipoHelper.ExpedientesParaCambioTipo_SelectionState("", "EXP-1", errMsg))
    If CBool(payload("enableSeleccionar")) <> False Then Err.Raise 1001, , "expected disabled"
    Test_ExpedientesParaCambioTipoHelper_SelectionState_EdgeEmpty = BuildOk("empty-selection", logs)
    Exit Function
EH:
    Test_ExpedientesParaCambioTipoHelper_SelectionState_EdgeEmpty = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedientesParaCambioTipoHelper_SelectedType_SadRegular() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, json As String
    json = modExpedientesParaCambioTipoHelper.ExpedientesParaCambioTipo_ValidateSelectedType("Regular", errMsg)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected non-AM/Lote rejection"
    Test_ExpedientesParaCambioTipoHelper_SelectedType_SadRegular = BuildOk("regular-rejected", logs)
    Exit Function
EH:
    Test_ExpedientesParaCambioTipoHelper_SelectedType_SadRegular = BuildFail(Err.Description, logs)
End Function
