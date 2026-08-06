Attribute VB_Name = "modExpedientesGestionTecnicaHelper"
Option Compare Database
Option Explicit

' Pure-data helpers for Form_FormExpedientesGestionTecnica.
' Public prefix: ExpedientesGestionTecnica_*

Private Function BuildJsonPayload(ByVal p_Ok As Boolean, ByVal p_Payload As Object, ByVal p_ErrorMsg As String, ByRef p_Logs() As String) As String
    Dim payloadJson As String
    If p_Payload Is Nothing Then payloadJson = "null" Else payloadJson = JsonConverter.ConvertToJson(p_Payload)
    Dim errorJson As String
    If p_Ok Then errorJson = "null" Else errorJson = """" & TestHelper.EscapeJsonString(p_ErrorMsg) & """"
    BuildJsonPayload = "{""ok"":" & LCase$(CStr(p_Ok)) & ",""value"":null,""payload"":" & payloadJson & ",""error"":" & errorJson & ",""logs"":" & TestHelper.JsonStringArray(p_Logs) & "}"
End Function

Private Function Envelope(ByVal p_Payload As Object, Optional ByVal p_ErrorMsg As String = "") As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    Envelope = BuildJsonPayload(Len(p_ErrorMsg) = 0, p_Payload, p_ErrorMsg, logs)
End Function

Public Function ExpedientesGestionTecnica_ResetFieldValue(ByVal p_FieldName As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("field") = p_FieldName
    payload("recognized") = True
    Select Case UCase$(Trim$(p_FieldName))
        Case "PALABRACLAVE"
            payload("value") = ""
        Case "ESTADO", "JURIDICA", "CODEXP", "JP"
            payload("value") = "Todos"
        Case Else
            payload("recognized") = False
            payload("value") = ""
    End Select
    payload("shouldFilter") = CBool(payload("recognized"))
    ExpedientesGestionTecnica_ResetFieldValue = Envelope(payload)
End Function

Public Function ExpedientesGestionTecnica_ListHeader(Optional ByRef p_Error As String) As String
    Dim payload As Object
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("rowSource") = "IDExpediente;Tipo;NºExp;Título;FInicio;FFin;FFinGarantía;Estado"
    payload("label") = "Lista de Expedientes Filtrados"
    ExpedientesGestionTecnica_ListHeader = Envelope(payload)
End Function

Public Function ExpedientesGestionTecnica_SearchChangedDecision(ByVal p_CurrentSignature As String, ByVal p_PreviousSignature As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("hasChanged") = (Trim$(p_CurrentSignature) <> Trim$(p_PreviousSignature))
    payload("shouldRefresh") = CBool(payload("hasChanged"))
    ExpedientesGestionTecnica_SearchChangedDecision = Envelope(payload)
End Function

Public Function ExpedientesGestionTecnica_SelectionState(ByVal p_SelectedID As String, ByVal p_CurrentID As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("hasSelection") = (Len(Trim$(p_SelectedID)) > 0)
    payload("requiresLoad") = (Len(Trim$(p_SelectedID)) > 0 And Trim$(p_SelectedID) <> Trim$(p_CurrentID))
    payload("enableDetail") = CBool(payload("hasSelection"))
    ExpedientesGestionTecnica_SelectionState = Envelope(payload)
End Function
