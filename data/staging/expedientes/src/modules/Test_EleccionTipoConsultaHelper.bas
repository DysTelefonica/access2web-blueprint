Attribute VB_Name = "Test_EleccionTipoConsultaHelper"
Option Compare Database
Option Explicit

' Test_EleccionTipoConsultaHelper — pure-data atoms for Phase 3.5 / PR42.

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

Public Function Test_EleccionTipoConsultaHelper_ParseOpenArgs_Defaults() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modEleccionTipoConsultaHelper.EleccionTipoConsulta_ParseOpenArgs("", "Todos", errMsg))
    If CStr(payload("nombreLista")) <> "Todos" Then Err.Raise 1001, , "expected Todos"
    If CStr(payload("nombreFormulario")) <> "FormExpedientesGestion" Then Err.Raise 1002, , "expected default form"
    Test_EleccionTipoConsultaHelper_ParseOpenArgs_Defaults = BuildOk("defaults", logs)
    Exit Function
EH:
    Test_EleccionTipoConsultaHelper_ParseOpenArgs_Defaults = BuildFail(Err.Description, logs)
End Function

Public Function Test_EleccionTipoConsultaHelper_ParseOpenArgs_RejectsInsufficient() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, json As String
    json = modEleccionTipoConsultaHelper.EleccionTipoConsulta_ParseOpenArgs("SoloUno", "Todos", errMsg)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected fail JSON"
    Test_EleccionTipoConsultaHelper_ParseOpenArgs_RejectsInsufficient = BuildOk("invalid", logs)
    Exit Function
EH:
    Test_EleccionTipoConsultaHelper_ParseOpenArgs_RejectsInsufficient = BuildFail(Err.Description, logs)
End Function

Public Function Test_EleccionTipoConsultaHelper_SelectType_SimpleSkipsLoadedCollections() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modEleccionTipoConsultaHelper.EleccionTipoConsulta_SelectType("Consulta Simple", True, errMsg))
    If CStr(payload("fieldSet")) <> "simple" Then Err.Raise 1001, , "expected simple"
    If CBool(payload("loadCollections")) <> False Then Err.Raise 1002, , "expected no load"
    Test_EleccionTipoConsultaHelper_SelectType_SimpleSkipsLoadedCollections = BuildOk("simple", logs)
    Exit Function
EH:
    Test_EleccionTipoConsultaHelper_SelectType_SimpleSkipsLoadedCollections = BuildFail(Err.Description, logs)
End Function

Public Function Test_EleccionTipoConsultaHelper_ListLoadPlan_TodosLoadsThreeLists() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modEleccionTipoConsultaHelper.EleccionTipoConsulta_ListLoadPlan("Todos", errMsg))
    If CBool(payload("loadAM")) <> True Then Err.Raise 1001, , "expected AM"
    If CBool(payload("loadLotes")) <> True Then Err.Raise 1002, , "expected lotes"
    If CBool(payload("loadBasados")) <> True Then Err.Raise 1003, , "expected basados"
    If CBool(payload("loadTecnica")) <> False Then Err.Raise 1004, , "expected no tecnica"
    Test_EleccionTipoConsultaHelper_ListLoadPlan_TodosLoadsThreeLists = BuildOk("todos", logs)
    Exit Function
EH:
    Test_EleccionTipoConsultaHelper_ListLoadPlan_TodosLoadsThreeLists = BuildFail(Err.Description, logs)
End Function
