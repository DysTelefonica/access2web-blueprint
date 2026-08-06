Attribute VB_Name = "Test_ExpedientesGestionHelper"
Option Compare Database
Option Explicit

' Test_ExpedientesGestionHelper — canonical JSON TDD atoms for Phase 3.4 / PR41.

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

Public Function Test_ExpedientesGestionHelper_FormLoadState_AdminShowsUltimoCambio() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modExpedientesGestionHelper.ExpedientesGestion_FormLoadState(True, True, "cambio", "Sí", errMsg))
    If CBool(payload("ultimaModificacionVisible")) <> True Then Err.Raise 1001, , "expected visible"
    If CStr(payload("ultimaModificacionCaption")) <> "cambio" Then Err.Raise 1002, , "expected caption"
    Test_ExpedientesGestionHelper_FormLoadState_AdminShowsUltimoCambio = BuildOk("form-state", logs)
    Exit Function
EH:
    Test_ExpedientesGestionHelper_FormLoadState_AdminShowsUltimoCambio = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedientesGestionHelper_FormLoadState_ReadOnlyHidesUltimoCambio() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modExpedientesGestionHelper.ExpedientesGestion_FormLoadState(False, True, "", "No", errMsg))
    If CBool(payload("ultimaModificacionVisible")) <> False Then Err.Raise 1001, , "expected hidden"
    Test_ExpedientesGestionHelper_FormLoadState_ReadOnlyHidesUltimoCambio = BuildOk("readonly-state", logs)
    Exit Function
EH:
    Test_ExpedientesGestionHelper_FormLoadState_ReadOnlyHidesUltimoCambio = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedientesGestionHelper_ResetEstado_DefaultTodos() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modExpedientesGestionHelper.ExpedientesGestion_ResetFieldValue("ESTADO", errMsg))
    If CStr(payload("value")) <> "Todos" Then Err.Raise 1001, , "expected Todos"
    Test_ExpedientesGestionHelper_ResetEstado_DefaultTodos = BuildOk("reset-estado", logs)
    Exit Function
EH:
    Test_ExpedientesGestionHelper_ResetEstado_DefaultTodos = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedientesGestionHelper_SelectionState_EmptyDisablesDetail() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modExpedientesGestionHelper.ExpedientesGestion_SelectionState("", "EXP-1", errMsg))
    If CBool(payload("enableDetail")) <> False Then Err.Raise 1001, , "expected detail disabled"
    Test_ExpedientesGestionHelper_SelectionState_EmptyDisablesDetail = BuildOk("empty-selection", logs)
    Exit Function
EH:
    Test_ExpedientesGestionHelper_SelectionState_EmptyDisablesDetail = BuildFail(Err.Description, logs)
End Function
