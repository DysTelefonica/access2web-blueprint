Attribute VB_Name = "modGradoClasificacionFormHelper"
Option Compare Database
Option Explicit

' Pure-data helpers for Form_FormGradoClasificacion (singular edit form).
' Existing modGradosClasificacionHelper belongs to Form_FormGradosClasificacionGestion.
' Public prefix: GradoClasificacionForm_*

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

Public Function GradoClasificacionForm_BuildValues(ByVal p_Grado As String, ByVal p_Descripcion As String, Optional ByRef p_Error As String) As Object
    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")
    values.CompareMode = TextCompare
    values("GradoClasificacion") = p_Grado
    values("DESCRIPCION") = p_Descripcion
    Set GradoClasificacionForm_BuildValues = values
End Function

Public Function GradoClasificacionForm_FormOpenState(ByVal p_EsAdministrador As Boolean, ByVal p_IsEdit As Boolean, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    payload("editarEnabled") = p_EsAdministrador
    If p_IsEdit Then payload("title") = "EDICIÓN DE GRADO CLASIFICACIÓN" Else payload("title") = "ALTA DE GRADO CLASIFICACIÓN"
    GradoClasificacionForm_FormOpenState = Envelope(payload)
End Function

Public Function GradoClasificacionForm_HasChanges(ByVal p_Initial As Object, ByVal p_CurrentValues As Object, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    If p_Initial Is Nothing Then
        payload("hasChanges") = True
    Else
        payload("hasChanges") = Helper_EntidadCRUD.HaHabidoCambiosGenerico(p_Initial, p_CurrentValues, Array("GradoClasificacion", "DESCRIPCION"), p_Error)
        If p_Error <> "" Then
            GradoClasificacionForm_HasChanges = Envelope(Nothing, p_Error)
            Exit Function
        End If
    End If
    GradoClasificacionForm_HasChanges = Envelope(payload)
End Function

Public Function GradoClasificacionForm_RegisterDecision(ByVal p_HasChanges As Boolean, ByVal p_IsNew As Boolean, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    If Not p_HasChanges Then
        p_Error = "No hay cambios que guardar"
        GradoClasificacionForm_RegisterDecision = Envelope(Nothing, p_Error)
        Exit Function
    End If
    If p_IsNew Then payload("eventName") = "Alta" Else payload("eventName") = "Editado"
    payload("shouldRegister") = True
    GradoClasificacionForm_RegisterDecision = Envelope(payload)
End Function
