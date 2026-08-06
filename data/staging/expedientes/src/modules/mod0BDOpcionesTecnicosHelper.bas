Attribute VB_Name = "mod0BDOpcionesTecnicosHelper"
Option Compare Database
Option Explicit

' Pure-data helpers for Form_Form0BDOpcionesTecnicos.
' Public prefix: OpcionesTecnicos_*

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

Public Function OpcionesTecnicos_CloseDecision(ByVal p_MainOptionsOpen As Boolean, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    If p_MainOptionsOpen Then payload("action") = "close-form" Else payload("action") = "close-database"
    OpcionesTecnicos_CloseDecision = Envelope(payload)
End Function

Public Function OpcionesTecnicos_SearchTarget(ByVal p_SearchKind As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    Select Case LCase$(Trim$(p_SearchKind))
        Case "completa"
            payload("formName") = "FormExpedientesGestion"
            payload("openArgs") = "SoloLectura"
        Case "simple"
            payload("formName") = "FormExpedientesGestionTecnica"
            payload("openArgs") = ""
        Case Else
            p_Error = "Tipo de búsqueda no reconocido"
            OpcionesTecnicos_SearchTarget = Envelope(Nothing, p_Error)
            Exit Function
    End Select
    payload("closeIfOpen") = True
    OpcionesTecnicos_SearchTarget = Envelope(payload)
End Function

Public Function OpcionesTecnicos_FormOpenState(ByVal p_Version As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    payload("versionCaption") = "Versión: " & p_Version
    payload("captionUsesTitle") = True
    OpcionesTecnicos_FormOpenState = Envelope(payload)
End Function

Public Function OpcionesTecnicos_LabelDispatch(ByVal p_Enabled As Boolean, ByVal p_SearchKind As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    If p_Enabled Then payload("searchKind") = p_SearchKind Else payload("searchKind") = "none"
    OpcionesTecnicos_LabelDispatch = Envelope(payload)
End Function
