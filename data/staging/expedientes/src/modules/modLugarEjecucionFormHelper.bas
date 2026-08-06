Attribute VB_Name = "modLugarEjecucionFormHelper"
Option Compare Database
Option Explicit

' Pure-data helpers for Form_FormLugarEjecucion (singular edit form).
' Existing modLugarEjecucionHelper belongs to Form_FormLugarEjecucionGestion.
' Public prefix: LugarEjecucionForm_*

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

Public Function LugarEjecucionForm_BuildValues(ByVal p_Lugar As String, ByVal p_Descripcion As String, Optional ByRef p_Error As String) As Object
    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")
    values.CompareMode = TextCompare
    values("LugarEjecucion") = p_Lugar
    values("DESCRIPCION") = p_Descripcion
    Set LugarEjecucionForm_BuildValues = values
End Function

Public Function LugarEjecucionForm_FormOpenState(ByVal p_EsAdministrador As Boolean, ByVal p_IsEdit As Boolean, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    payload("registrarEnabled") = p_EsAdministrador
    If p_IsEdit Then payload("title") = "EDICIÓN DE LUGAR EJECUCIÓN" Else payload("title") = "ALTA DE LUGAR EJECUCIÓN"
    LugarEjecucionForm_FormOpenState = Envelope(payload)
End Function

Public Function LugarEjecucionForm_HasChanges(ByVal p_Initial As Object, ByVal p_CurrentValues As Object, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    If p_Initial Is Nothing Then
        payload("hasChanges") = True
    Else
        payload("hasChanges") = Helper_EntidadCRUD.HaHabidoCambiosGenerico(p_Initial, p_CurrentValues, Array("LugarEjecucion", "DESCRIPCION"), p_Error)
        If p_Error <> "" Then
            LugarEjecucionForm_HasChanges = Envelope(Nothing, p_Error)
            Exit Function
        End If
    End If
    LugarEjecucionForm_HasChanges = Envelope(payload)
End Function

Public Function LugarEjecucionForm_RegisterDecision(ByVal p_HasChanges As Boolean, ByVal p_IsNew As Boolean, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    If Not p_HasChanges Then
        p_Error = "No hay cambios que guardar"
        LugarEjecucionForm_RegisterDecision = Envelope(Nothing, p_Error)
        Exit Function
    End If
    If p_IsNew Then payload("eventName") = "Alta" Else payload("eventName") = "Editado"
    payload("shouldRegister") = True
    LugarEjecucionForm_RegisterDecision = Envelope(payload)
End Function
