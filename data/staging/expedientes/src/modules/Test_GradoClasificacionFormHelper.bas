Attribute VB_Name = "Test_GradoClasificacionFormHelper"
Option Compare Database
Option Explicit

' Test_GradoClasificacionFormHelper — pure-data atoms for Phase 3.5 / PR42.

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

Public Function Test_GradoClasificacionFormHelper_FormOpenState_NonAdminDisablesEdit() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modGradoClasificacionFormHelper.GradoClasificacionForm_FormOpenState(False, False, errMsg))
    If CBool(payload("editarEnabled")) <> False Then Err.Raise 1001, , "expected disabled"
    If CStr(payload("title")) <> "ALTA DE GRADO CLASIFICACIÓN" Then Err.Raise 1002, , "expected alta title"
    Test_GradoClasificacionFormHelper_FormOpenState_NonAdminDisablesEdit = BuildOk("state", logs)
    Exit Function
EH:
    Test_GradoClasificacionFormHelper_FormOpenState_NonAdminDisablesEdit = BuildFail(Err.Description, logs)
End Function

Public Function Test_GradoClasificacionFormHelper_BuildValues_MapsFields() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, values As Object
    Set values = modGradoClasificacionFormHelper.GradoClasificacionForm_BuildValues("Reservado", "Desc", errMsg)
    If CStr(values("GradoClasificacion")) <> "Reservado" Then Err.Raise 1001, , "expected GradoClasificacion"
    If CStr(values("DESCRIPCION")) <> "Desc" Then Err.Raise 1002, , "expected desc"
    Test_GradoClasificacionFormHelper_BuildValues_MapsFields = BuildOk("values", logs)
    Exit Function
EH:
    Test_GradoClasificacionFormHelper_BuildValues_MapsFields = BuildFail(Err.Description, logs)
End Function

Public Function Test_GradoClasificacionFormHelper_RegisterDecision_NoChangesRejected() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, json As String
    json = modGradoClasificacionFormHelper.GradoClasificacionForm_RegisterDecision(False, False, errMsg)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected fail JSON"
    Test_GradoClasificacionFormHelper_RegisterDecision_NoChangesRejected = BuildOk("no-changes", logs)
    Exit Function
EH:
    Test_GradoClasificacionFormHelper_RegisterDecision_NoChangesRejected = BuildFail(Err.Description, logs)
End Function

Public Function Test_GradoClasificacionFormHelper_RegisterDecision_EditRaisesEditado() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(modGradoClasificacionFormHelper.GradoClasificacionForm_RegisterDecision(True, False, errMsg))
    If CStr(payload("eventName")) <> "Editado" Then Err.Raise 1001, , "expected Editado"
    Test_GradoClasificacionFormHelper_RegisterDecision_EditRaisesEditado = BuildOk("editado", logs)
    Exit Function
EH:
    Test_GradoClasificacionFormHelper_RegisterDecision_EditRaisesEditado = BuildFail(Err.Description, logs)
End Function
