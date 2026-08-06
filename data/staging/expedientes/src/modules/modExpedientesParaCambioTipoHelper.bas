Attribute VB_Name = "modExpedientesParaCambioTipoHelper"
Option Compare Database
Option Explicit

' Pure-data helpers for Form_FormExpedientesParaCambioTipo.
' Public prefix: ExpedientesParaCambioTipo_*

Private Function BuildJsonPayload(ByVal p_Ok As Boolean, ByVal p_Payload As Object, ByVal p_ErrorMsg As String, ByRef p_Logs() As String) As String
    Dim payloadJson As String
    If p_Payload Is Nothing Then payloadJson = "null" Else payloadJson = JsonConverter.ConvertToJson(p_Payload)
    Dim errorJson As String
    If p_Ok Then errorJson = "null" Else errorJson = """" & TestHelper.EscapeJsonString(p_ErrorMsg) & """"
    BuildJsonPayload = "{""ok"":" & LCase$(CStr(p_Ok)) & ",""value"":null,""payload"":" & payloadJson & ",""error"":" & errorJson & ",""logs"":" & TestHelper.JsonStringArray(p_Logs) & "}"
End Function

Private Function PayloadEnvelope(ByVal p_Payload As Object, Optional ByVal p_ErrorMsg As String = "") As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    PayloadEnvelope = BuildJsonPayload(Len(p_ErrorMsg) = 0, p_Payload, p_ErrorMsg, logs)
End Function

Public Function ExpedientesParaCambioTipo_ListHeader(Optional ByRef p_Error As String) As String
    Dim payload As Object
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("rowSource") = "IDExpediente;Tipo;NºExp;Título;FInicio;FFin;FFinGarantía"
    ExpedientesParaCambioTipo_ListHeader = PayloadEnvelope(payload)
End Function

Public Function ExpedientesParaCambioTipo_ValidateSearchTerm(ByVal p_PalabraClave As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("canSearch") = (Len(Trim$(p_PalabraClave)) > 0)
    payload("normalized") = Trim$(p_PalabraClave)
    If Not CBool(payload("canSearch")) Then p_Error = "PalabraClave is required"
    ExpedientesParaCambioTipo_ValidateSearchTerm = BuildJsonPayload(CBool(payload("canSearch")), payload, p_Error, logs)
End Function

Public Function ExpedientesParaCambioTipo_SelectionState(ByVal p_SelectedID As String, ByVal p_CurrentID As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("hasSelection") = (Len(Trim$(p_SelectedID)) > 0)
    payload("requiresLoad") = (Len(Trim$(p_SelectedID)) > 0 And Trim$(p_SelectedID) <> Trim$(p_CurrentID))
    payload("enableSeleccionar") = CBool(payload("hasSelection"))
    ExpedientesParaCambioTipo_SelectionState = PayloadEnvelope(payload)
End Function

Public Function ExpedientesParaCambioTipo_ValidateSelectedType(ByVal p_TipoTexto As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Dim normalized As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""
    normalized = UCase$(Trim$(p_TipoTexto))
    Set payload = CreateObject("Scripting.Dictionary")
    payload("canSelect") = (normalized = "AM" Or normalized = "LOTE")
    payload("tipo") = normalized
    If Not CBool(payload("canSelect")) Then p_Error = "Se ha de seleccionar un AM o un Lote"
    ExpedientesParaCambioTipo_ValidateSelectedType = BuildJsonPayload(CBool(payload("canSelect")), payload, p_Error, logs)
End Function
