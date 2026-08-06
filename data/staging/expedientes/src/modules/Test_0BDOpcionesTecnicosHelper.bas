Attribute VB_Name = "Test_0BDOpcionesTecnicosHelper"
Option Compare Database
Option Explicit

' Test_0BDOpcionesTecnicosHelper — pure-data atoms for Phase 3.5 / PR42.

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

Public Function Test_0BDOpcionesTecnicosHelper_CloseDecision_MainOpenClosesForm() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(mod0BDOpcionesTecnicosHelper.OpcionesTecnicos_CloseDecision(True, errMsg))
    If CStr(payload("action")) <> "close-form" Then Err.Raise 1001, , "expected close-form"
    Test_0BDOpcionesTecnicosHelper_CloseDecision_MainOpenClosesForm = BuildOk("close", logs)
    Exit Function
EH:
    Test_0BDOpcionesTecnicosHelper_CloseDecision_MainOpenClosesForm = BuildFail(Err.Description, logs)
End Function

Public Function Test_0BDOpcionesTecnicosHelper_SearchTarget_CompletaUsesReadOnlyGestion() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(mod0BDOpcionesTecnicosHelper.OpcionesTecnicos_SearchTarget("completa", errMsg))
    If CStr(payload("formName")) <> "FormExpedientesGestion" Then Err.Raise 1001, , "expected gestion"
    If CStr(payload("openArgs")) <> "SoloLectura" Then Err.Raise 1002, , "expected SoloLectura"
    Test_0BDOpcionesTecnicosHelper_SearchTarget_CompletaUsesReadOnlyGestion = BuildOk("completa", logs)
    Exit Function
EH:
    Test_0BDOpcionesTecnicosHelper_SearchTarget_CompletaUsesReadOnlyGestion = BuildFail(Err.Description, logs)
End Function

Public Function Test_0BDOpcionesTecnicosHelper_SearchTarget_SimpleUsesTecnica() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, payload As Object
    Set payload = PayloadOf(mod0BDOpcionesTecnicosHelper.OpcionesTecnicos_SearchTarget("simple", errMsg))
    If CStr(payload("formName")) <> "FormExpedientesGestionTecnica" Then Err.Raise 1001, , "expected tecnica"
    Test_0BDOpcionesTecnicosHelper_SearchTarget_SimpleUsesTecnica = BuildOk("simple", logs)
    Exit Function
EH:
    Test_0BDOpcionesTecnicosHelper_SearchTarget_SimpleUsesTecnica = BuildFail(Err.Description, logs)
End Function

Public Function Test_0BDOpcionesTecnicosHelper_SearchTarget_RejectsUnknown() As String
    Dim logs() As String: logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim errMsg As String, json As String
    json = mod0BDOpcionesTecnicosHelper.OpcionesTecnicos_SearchTarget("x", errMsg)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected fail JSON"
    Test_0BDOpcionesTecnicosHelper_SearchTarget_RejectsUnknown = BuildOk("unknown", logs)
    Exit Function
EH:
    Test_0BDOpcionesTecnicosHelper_SearchTarget_RejectsUnknown = BuildFail(Err.Description, logs)
End Function
