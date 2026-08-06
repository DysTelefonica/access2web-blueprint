Attribute VB_Name = "modExpedienteGeneralHelper"
Option Compare Database
Option Explicit

' Pure-data helpers for Form_FormExpedienteGeneral.
' Public prefix: ExpedienteGeneral_*

Private Function BuildJsonPayload(ByVal p_Ok As Boolean, ByVal p_Payload As Object, ByVal p_ErrorMsg As String, ByRef p_Logs() As String) As String
    Dim payloadJson As String
    If p_Payload Is Nothing Then payloadJson = "null" Else payloadJson = JsonConverter.ConvertToJson(p_Payload)
    Dim errorJson As String
    If p_Ok Then errorJson = "null" Else errorJson = """" & TestHelper.EscapeJsonString(p_ErrorMsg) & """"
    BuildJsonPayload = "{""ok"":" & LCase$(CStr(p_Ok)) & ",""value"":null,""payload"":" & payloadJson & ",""error"":" & errorJson & ",""logs"":" & TestHelper.JsonStringArray(p_Logs) & "}"
End Function

Public Function ExpedienteGeneral_FormLoadState(ByVal p_AbiertoParaEditar As Boolean, ByVal p_EsAdministrador As Boolean, Optional ByRef p_Error As String) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""
    Dim allowEdits As Boolean
    allowEdits = (p_AbiertoParaEditar And p_EsAdministrador)
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("allowEdits") = allowEdits
    payload("parentRegistrarVisible") = True
    payload("parentRegistrarEnabled") = allowEdits
    payload("pickerButtonsEnabled") = allowEdits
    ExpedienteGeneral_FormLoadState = BuildJsonPayload(True, payload, "", logs)
End Function

Public Function ExpedienteGeneral_AmbitoAfterUpdate(ByVal p_Ambito As Variant, Optional ByRef p_Error As String) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("setHPSAplica") = (Nz(p_Ambito, "") = "HPS")
    payload("hpsAplica") = "Sí"
    ExpedienteGeneral_AmbitoAfterUpdate = BuildJsonPayload(True, payload, "", logs)
End Function

Public Function ExpedienteGeneral_AccesoSharepointState(ByVal p_AccesoSharepoint As Variant, Optional ByRef p_Error As String) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(1)
    p_Error = ""
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("verSharepointEnabled") = (Nz(p_AccesoSharepoint, "") <> "")
    ExpedienteGeneral_AccesoSharepointState = BuildJsonPayload(True, payload, "", logs)
End Function

Public Function ExpedienteGeneral_UnloadDecision(ByVal p_AllowEdits As Boolean, ByVal p_SnapshotInicial As String, ByVal p_SnapshotActual As String, Optional ByRef p_Error As String) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("guardar") = (p_AllowEdits And p_SnapshotInicial <> p_SnapshotActual)
    ExpedienteGeneral_UnloadDecision = BuildJsonPayload(True, payload, "", logs)
End Function

Public Function ExpedienteGeneral_TextoOrdinalUsado(ByVal p_CodExp As String, ByVal p_Nemotecnico As String, ByVal p_Titulo As String, Optional ByRef p_Error As String) As String
    p_Error = ""
    If p_CodExp <> "" Then
        ExpedienteGeneral_TextoOrdinalUsado = p_CodExp
    ElseIf p_Nemotecnico <> "" Then
        ExpedienteGeneral_TextoOrdinalUsado = p_Nemotecnico
    Else
        ExpedienteGeneral_TextoOrdinalUsado = p_Titulo
    End If
End Function
