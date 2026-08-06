Attribute VB_Name = "Test_ExpedientesGestionTecnicaHelper"
Option Compare Database
Option Explicit

' Test_ExpedientesGestionTecnicaHelper — canonical JSON TDD atoms for Phase 3.4 / PR41.

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

Public Function Test_ExpedientesGestionTecnicaHelper_ResetPalabraClave_BlanksValue() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modExpedientesGestionTecnicaHelper.ExpedientesGestionTecnica_ResetFieldValue("PalabraClave", errMsg))
    If CStr(payload("value")) <> "" Then Err.Raise 1001, , "expected blank value"
    Test_ExpedientesGestionTecnicaHelper_ResetPalabraClave_BlanksValue = BuildOk("reset-palabra", logs)
    Exit Function
EH:
    Test_ExpedientesGestionTecnicaHelper_ResetPalabraClave_BlanksValue = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedientesGestionTecnicaHelper_ResetEstado_DefaultTodos() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modExpedientesGestionTecnicaHelper.ExpedientesGestionTecnica_ResetFieldValue("ESTADO", errMsg))
    If CStr(payload("value")) <> "Todos" Then Err.Raise 1001, , "expected Todos"
    Test_ExpedientesGestionTecnicaHelper_ResetEstado_DefaultTodos = BuildOk("reset-estado", logs)
    Exit Function
EH:
    Test_ExpedientesGestionTecnicaHelper_ResetEstado_DefaultTodos = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedientesGestionTecnicaHelper_SearchChanged_SameSignatureSkips() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modExpedientesGestionTecnicaHelper.ExpedientesGestionTecnica_SearchChangedDecision("a|b", "a|b", errMsg))
    If CBool(payload("shouldRefresh")) <> False Then Err.Raise 1001, , "expected shouldRefresh=false"
    Test_ExpedientesGestionTecnicaHelper_SearchChanged_SameSignatureSkips = BuildOk("same-signature", logs)
    Exit Function
EH:
    Test_ExpedientesGestionTecnicaHelper_SearchChanged_SameSignatureSkips = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedientesGestionTecnicaHelper_Selection_EnablesDetail() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modExpedientesGestionTecnicaHelper.ExpedientesGestionTecnica_SelectionState("EXP-1", "", errMsg))
    If CBool(payload("enableDetail")) <> True Then Err.Raise 1001, , "expected detail enabled"
    Test_ExpedientesGestionTecnicaHelper_Selection_EnablesDetail = BuildOk("selection", logs)
    Exit Function
EH:
    Test_ExpedientesGestionTecnicaHelper_Selection_EnablesDetail = BuildFail(Err.Description, logs)
End Function
