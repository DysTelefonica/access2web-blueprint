Attribute VB_Name = "Test_ExpedienteSuministradoresHelper"
Option Compare Database
Option Explicit

' Test_ExpedienteSuministradoresHelper — canonical JSON TDD atoms for PR #40.

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

Public Function Test_ExpedienteSuministradoresHelper_ExtraerTag_HappyRELID() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim relId As String
    relId = modExpedienteSuministradoresHelper.ExpedienteSuministradores_ExtraerTag("RELID=123;IDS=456", "RELID", p_Error)
    If relId <> "123" Then Err.Raise 1001, , "expected RELID=123"
    Test_ExpedienteSuministradoresHelper_ExtraerTag_HappyRELID = BuildOk("relid", logs)
    Exit Function
EH:
    Test_ExpedienteSuministradoresHelper_ExtraerTag_HappyRELID = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteSuministradoresHelper_ExtraerTag_EdgeMissing() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim value As String
    value = modExpedienteSuministradoresHelper.ExpedienteSuministradores_ExtraerTag("RELID=123", "IDS", p_Error)
    If value <> "" Then Err.Raise 1001, , "expected missing tag empty"
    Test_ExpedienteSuministradoresHelper_ExtraerTag_EdgeMissing = BuildOk("missing-empty", logs)
    Exit Function
EH:
    Test_ExpedienteSuministradoresHelper_ExtraerTag_EdgeMissing = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteSuministradoresHelper_GestionarBotones_HappyChildEditable() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteSuministradoresHelper.ExpedienteSuministradores_GestionarBotonesState(True, "K123", True, p_Error))
    If CBool(payload("eliminarEnabled")) <> True Then Err.Raise 1001, , "expected eliminarEnabled=true"
    If CBool(payload("verDetalleEnabled")) <> True Then Err.Raise 1002, , "expected verDetalleEnabled=true"
    Test_ExpedienteSuministradoresHelper_GestionarBotones_HappyChildEditable = BuildOk("child-editable", logs)
    Exit Function
EH:
    Test_ExpedienteSuministradoresHelper_GestionarBotones_HappyChildEditable = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteSuministradoresHelper_GestionarBotones_SadRoot() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteSuministradoresHelper.ExpedienteSuministradores_GestionarBotonesState(True, "ROOT_ORGANO", True, p_Error))
    If CBool(payload("eliminarEnabled")) <> False Then Err.Raise 1001, , "expected eliminarEnabled=false"
    If CBool(payload("verDetalleEnabled")) <> False Then Err.Raise 1002, , "expected verDetalleEnabled=false"
    Test_ExpedienteSuministradoresHelper_GestionarBotones_SadRoot = BuildOk("root-disabled", logs)
    Exit Function
EH:
    Test_ExpedienteSuministradoresHelper_GestionarBotones_SadRoot = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteSuministradoresHelper_ResolveDropTarget_HappyRoot() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteSuministradoresHelper.ExpedienteSuministradores_ResolveDropTarget("ROOT_ORGANO", "", False, p_Error))
    If CStr(payload("tipo")) <> "ROOT_CONTR" Then Err.Raise 1001, , "expected ROOT_CONTR"
    If CStr(payload("keyDest")) <> "ROOT_ORGANO" Then Err.Raise 1002, , "expected ROOT_ORGANO keyDest"
    Test_ExpedienteSuministradoresHelper_ResolveDropTarget_HappyRoot = BuildOk("root-target", logs)
    Exit Function
EH:
    Test_ExpedienteSuministradoresHelper_ResolveDropTarget_HappyRoot = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteSuministradoresHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH
    logs(0) = Test_ExpedienteSuministradoresHelper_ExtraerTag_HappyRELID()
    logs(1) = Test_ExpedienteSuministradoresHelper_ExtraerTag_EdgeMissing()
    logs(2) = Test_ExpedienteSuministradoresHelper_GestionarBotones_HappyChildEditable()
    logs(3) = Test_ExpedienteSuministradoresHelper_GestionarBotones_SadRoot()
    logs(4) = Test_ExpedienteSuministradoresHelper_ResolveDropTarget_HappyRoot()
    Test_ExpedienteSuministradoresHelper_RunAll = BuildOk("5 atoms executed", logs)
    Exit Function
EH:
    Test_ExpedienteSuministradoresHelper_RunAll = BuildFail(Err.Description, logs)
End Function
