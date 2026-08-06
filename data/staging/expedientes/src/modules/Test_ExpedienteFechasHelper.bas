Attribute VB_Name = "Test_ExpedienteFechasHelper"
Option Compare Database
Option Explicit

' Test_ExpedienteFechasHelper — canonical JSON TDD atoms for PR #40.

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

Public Function Test_ExpedienteFechasHelper_Calcular_HappyCertificacion() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim fechaFin As Variant
    Dim p_Error As String
    Call modExpedienteFechasHelper.ExpedienteFechas_CalcularFechaFinGarantia(12, #1/15/2026#, #1/1/2026#, fechaFin, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If Not IsDate(fechaFin) Then Err.Raise 1001, , "expected date result"
    Test_ExpedienteFechasHelper_Calcular_HappyCertificacion = BuildOk(CStr(fechaFin), logs)
    Exit Function
EH:
    Test_ExpedienteFechasHelper_Calcular_HappyCertificacion = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteFechasHelper_Calcular_EdgeNoGarantia() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim fechaFin As Variant
    Dim p_Error As String
    Call modExpedienteFechasHelper.ExpedienteFechas_CalcularFechaFinGarantia("", #1/15/2026#, #1/1/2026#, fechaFin, p_Error)
    If Not IsNull(fechaFin) Then Err.Raise 1001, , "expected Null result"
    Test_ExpedienteFechasHelper_Calcular_EdgeNoGarantia = BuildOk("null-no-garantia", logs)
    Exit Function
EH:
    Test_ExpedienteFechasHelper_Calcular_EdgeNoGarantia = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteFechasHelper_FormLoadState_HappyAdminEditable() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteFechasHelper.ExpedienteFechas_FormLoadState(True, True, p_Error))
    If CBool(payload("allowEdits")) <> True Then Err.Raise 1001, , "expected allowEdits=true"
    If CBool(payload("perdidaEnabled")) <> False Then Err.Raise 1002, , "expected perdidaEnabled=false"
    Test_ExpedienteFechasHelper_FormLoadState_HappyAdminEditable = BuildOk("admin-edit", logs)
    Exit Function
EH:
    Test_ExpedienteFechasHelper_FormLoadState_HappyAdminEditable = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteFechasHelper_UnloadDecision_HappyChanged() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteFechasHelper.ExpedienteFechas_UnloadDecision(True, "a", "b", p_Error))
    If CBool(payload("guardar")) <> True Then Err.Raise 1001, , "expected guardar=true"
    Test_ExpedienteFechasHelper_UnloadDecision_HappyChanged = BuildOk("save-changed", logs)
    Exit Function
EH:
    Test_ExpedienteFechasHelper_UnloadDecision_HappyChanged = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteFechasHelper_UnloadDecision_EdgeReadOnly() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteFechasHelper.ExpedienteFechas_UnloadDecision(False, "a", "b", p_Error))
    If CBool(payload("guardar")) <> False Then Err.Raise 1001, , "expected guardar=false"
    Test_ExpedienteFechasHelper_UnloadDecision_EdgeReadOnly = BuildOk("readonly-no-save", logs)
    Exit Function
EH:
    Test_ExpedienteFechasHelper_UnloadDecision_EdgeReadOnly = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteFechasHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH
    logs(0) = Test_ExpedienteFechasHelper_Calcular_HappyCertificacion()
    logs(1) = Test_ExpedienteFechasHelper_Calcular_EdgeNoGarantia()
    logs(2) = Test_ExpedienteFechasHelper_FormLoadState_HappyAdminEditable()
    logs(3) = Test_ExpedienteFechasHelper_UnloadDecision_HappyChanged()
    logs(4) = Test_ExpedienteFechasHelper_UnloadDecision_EdgeReadOnly()
    Test_ExpedienteFechasHelper_RunAll = BuildOk("5 atoms executed", logs)
    Exit Function
EH:
    Test_ExpedienteFechasHelper_RunAll = BuildFail(Err.Description, logs)
End Function
