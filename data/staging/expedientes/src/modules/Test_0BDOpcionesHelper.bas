Attribute VB_Name = "Test_0BDOpcionesHelper"
Option Compare Database
Option Explicit

' Test_0BDOpcionesHelper — pure-data atoms for Phase 3.5 / PR42.

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

Public Function Test_0BDOpcionesHelper_FormOpenState_AdminEnablesControls() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(mod0BDOpcionesHelper.Opciones_FormOpenState(True, "1.2.3", errMsg))
    If CBool(payload("altaEnabled")) <> True Then Err.Raise 1001, , "expected alta enabled"
    If CBool(payload("buscadorTecnicoVisible")) <> True Then Err.Raise 1002, , "expected visible"
    If CStr(payload("versionCaption")) <> "Versión: 1.2.3" Then Err.Raise 1003, , "expected version"
    Test_0BDOpcionesHelper_FormOpenState_AdminEnablesControls = BuildOk("admin", logs)
    Exit Function
EH:
    Test_0BDOpcionesHelper_FormOpenState_AdminEnablesControls = BuildFail(Err.Description, logs)
End Function

Public Function Test_0BDOpcionesHelper_FormOpenState_NonAdminHidesTechnicalSearch() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(mod0BDOpcionesHelper.Opciones_FormOpenState(False, "1", errMsg))
    If CBool(payload("altaEnabled")) <> False Then Err.Raise 1001, , "expected alta disabled"
    If CBool(payload("buscadorTecnicoVisible")) <> False Then Err.Raise 1002, , "expected hidden"
    Test_0BDOpcionesHelper_FormOpenState_NonAdminHidesTechnicalSearch = BuildOk("non-admin", logs)
    Exit Function
EH:
    Test_0BDOpcionesHelper_FormOpenState_NonAdminHidesTechnicalSearch = BuildFail(Err.Description, logs)
End Function

Public Function Test_0BDOpcionesHelper_GestionTarget_TechnicianUsesTechnicalForm() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(mod0BDOpcionesHelper.Opciones_GestionExpedientesTarget(True, errMsg))
    If CStr(payload("formName")) <> "FormExpedientesGestionTecnica" Then Err.Raise 1001, , "expected technical form"
    Test_0BDOpcionesHelper_GestionTarget_TechnicianUsesTechnicalForm = BuildOk("target", logs)
    Exit Function
EH:
    Test_0BDOpcionesHelper_GestionTarget_TechnicianUsesTechnicalForm = BuildFail(Err.Description, logs)
End Function

Public Function Test_0BDOpcionesHelper_OpenAction_RejectsUnknown() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, json As String
    json = mod0BDOpcionesHelper.Opciones_OpenFormAction("unknown", errMsg)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected fail JSON"
    Test_0BDOpcionesHelper_OpenAction_RejectsUnknown = BuildOk("unknown", logs)
    Exit Function
EH:
    Test_0BDOpcionesHelper_OpenAction_RejectsUnknown = BuildFail(Err.Description, logs)
End Function
