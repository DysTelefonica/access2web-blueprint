Attribute VB_Name = "Test_ExpedienteGeneralHelper"
Option Compare Database
Option Explicit

' Test_ExpedienteGeneralHelper — canonical JSON TDD atoms for PR #40.

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

Public Function Test_ExpedienteGeneralHelper_FormLoadState_HappyAdminEditable() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteGeneralHelper.ExpedienteGeneral_FormLoadState(True, True, p_Error))
    If CBool(payload("allowEdits")) <> True Then Err.Raise 1001, , "expected allowEdits=true"
    If CBool(payload("parentRegistrarVisible")) <> True Then Err.Raise 1002, , "expected parentRegistrarVisible=true"
    Test_ExpedienteGeneralHelper_FormLoadState_HappyAdminEditable = BuildOk("admin-edit", logs)
    Exit Function
EH:
    Test_ExpedienteGeneralHelper_FormLoadState_HappyAdminEditable = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteGeneralHelper_AmbitoAfterUpdate_HappyHPS() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteGeneralHelper.ExpedienteGeneral_AmbitoAfterUpdate("HPS", p_Error))
    If CBool(payload("setHPSAplica")) <> True Then Err.Raise 1001, , "expected setHPSAplica=true"
    If CStr(payload("hpsAplica")) <> "Sí" Then Err.Raise 1002, , "expected hpsAplica=Sí"
    Test_ExpedienteGeneralHelper_AmbitoAfterUpdate_HappyHPS = BuildOk("hps", logs)
    Exit Function
EH:
    Test_ExpedienteGeneralHelper_AmbitoAfterUpdate_HappyHPS = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteGeneralHelper_AccesoSharepoint_EdgeEmpty() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteGeneralHelper.ExpedienteGeneral_AccesoSharepointState("", p_Error))
    If CBool(payload("verSharepointEnabled")) <> False Then Err.Raise 1001, , "expected button disabled"
    Test_ExpedienteGeneralHelper_AccesoSharepoint_EdgeEmpty = BuildOk("empty-disabled", logs)
    Exit Function
EH:
    Test_ExpedienteGeneralHelper_AccesoSharepoint_EdgeEmpty = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteGeneralHelper_UnloadDecision_HappyChanged() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteGeneralHelper.ExpedienteGeneral_UnloadDecision(True, "a", "b", p_Error))
    If CBool(payload("guardar")) <> True Then Err.Raise 1001, , "expected guardar=true"
    Test_ExpedienteGeneralHelper_UnloadDecision_HappyChanged = BuildOk("save-changed", logs)
    Exit Function
EH:
    Test_ExpedienteGeneralHelper_UnloadDecision_HappyChanged = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteGeneralHelper_TextoOrdinal_EdgeFallbackTitulo() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim value As String
    value = modExpedienteGeneralHelper.ExpedienteGeneral_TextoOrdinalUsado("", "", "titulo", p_Error)
    If value <> "titulo" Then Err.Raise 1001, , "expected titulo fallback"
    Test_ExpedienteGeneralHelper_TextoOrdinal_EdgeFallbackTitulo = BuildOk("titulo", logs)
    Exit Function
EH:
    Test_ExpedienteGeneralHelper_TextoOrdinal_EdgeFallbackTitulo = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteGeneralHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH
    logs(0) = Test_ExpedienteGeneralHelper_FormLoadState_HappyAdminEditable()
    logs(1) = Test_ExpedienteGeneralHelper_AmbitoAfterUpdate_HappyHPS()
    logs(2) = Test_ExpedienteGeneralHelper_AccesoSharepoint_EdgeEmpty()
    logs(3) = Test_ExpedienteGeneralHelper_UnloadDecision_HappyChanged()
    logs(4) = Test_ExpedienteGeneralHelper_TextoOrdinal_EdgeFallbackTitulo()
    Test_ExpedienteGeneralHelper_RunAll = BuildOk("5 atoms executed", logs)
    Exit Function
EH:
    Test_ExpedienteGeneralHelper_RunAll = BuildFail(Err.Description, logs)
End Function
