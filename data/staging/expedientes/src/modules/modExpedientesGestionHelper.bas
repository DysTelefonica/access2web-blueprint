Attribute VB_Name = "modExpedientesGestionHelper"
Option Compare Database
Option Explicit

' Pure-data helpers for Form_FormExpedientesGestion.
' Public prefix: ExpedientesGestion_*

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

Public Function ExpedientesGestion_FormLoadState(ByVal p_AbiertoParaEditar As Boolean, ByVal p_EsAdministrador As Boolean, ByVal p_UltimoCambioTexto As String, ByVal p_MostrarEstadoUsuario As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("ultimaModificacionVisible") = (p_AbiertoParaEditar And p_EsAdministrador)
    If Len(Trim$(p_UltimoCambioTexto)) = 0 Then
        payload("ultimaModificacionCaption") = "Aún sin registrar"
    Else
        payload("ultimaModificacionCaption") = p_UltimoCambioTexto
    End If
    payload("verEstado") = p_MostrarEstadoUsuario
    ExpedientesGestion_FormLoadState = Envelope(payload)
End Function

Public Function ExpedientesGestion_ResetFieldValue(ByVal p_FieldName As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("field") = p_FieldName
    payload("recognized") = True
    Select Case UCase$(Trim$(p_FieldName))
        Case "PALABRACLAVE"
            payload("value") = ""
        Case "ESTADO", "SUMINISTRADOR", "CODEXP", "COMERCIAL", "JP", "RAC", "IDEXPEDIENTE"
            payload("value") = "Todos"
        Case Else
            payload("recognized") = False
            payload("value") = ""
    End Select
    payload("shouldFilter") = CBool(payload("recognized"))
    ExpedientesGestion_ResetFieldValue = Envelope(payload)
End Function

Public Function ExpedientesGestion_SelectionState(ByVal p_SelectedID As String, ByVal p_CurrentID As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("hasSelection") = (Len(Trim$(p_SelectedID)) > 0)
    payload("requiresLoad") = (Len(Trim$(p_SelectedID)) > 0 And Trim$(p_SelectedID) <> Trim$(p_CurrentID))
    payload("enableDetail") = CBool(payload("hasSelection"))
    ExpedientesGestion_SelectionState = Envelope(payload)
End Function

Public Function ExpedientesGestion_ListVisualState(ByVal p_ShowEstado As Boolean, ByVal p_ListKind As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("listKind") = p_ListKind
    payload("showEstado") = p_ShowEstado
    If p_ShowEstado Then
        payload("amHeader") = "IDExpediente;Tipo;Nemotécnico;NºExp;Estado"
        payload("loteHeader") = "IDExpediente;Nº;Nemotécnico;Estado"
    Else
        payload("amHeader") = "IDExpediente;Tipo;Nemotécnico;NºExp;FInicial;FFinal"
        payload("loteHeader") = "IDExpediente;Nº;Nemotécnico;FInicial;FFinal"
    End If
    ExpedientesGestion_ListVisualState = Envelope(payload)
End Function
