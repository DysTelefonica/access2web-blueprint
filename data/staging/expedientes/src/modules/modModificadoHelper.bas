Attribute VB_Name = "modModificadoHelper"
Option Compare Database
Option Explicit

' Pure-data helpers for Form_FormModificado.
' Public prefix: Modificado_*

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

Private Function DictValue(ByVal p_Dict As Object, ByVal p_Key As String) As String
    If p_Dict Is Nothing Then DictValue = "": Exit Function
    If Not p_Dict.Exists(p_Key) Then DictValue = "": Exit Function
    DictValue = "" & p_Dict(p_Key)
End Function

Public Function Modificado_FormOpenState(ByVal p_EsAdministrador As Boolean, Optional ByRef p_Error As String) As String
    Dim payload As Object
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("allowEdits") = p_EsAdministrador
    payload("registrarEnabled") = p_EsAdministrador
    Modificado_FormOpenState = Envelope(payload)
End Function

Public Function Modificado_BuildFormValues(ByVal p_Descripcion As String, ByVal p_NModificado As String, ByVal p_FechaFirma As String, ByVal p_FechaFin As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("DESCRIPCION") = p_Descripcion
    payload("NModificado") = p_NModificado
    payload("FechaFirmaModificado") = p_FechaFirma
    payload("FechaFinModificado") = p_FechaFin
    Modificado_BuildFormValues = Envelope(payload)
End Function

Public Function Modificado_HasChanges(ByVal p_CurrentValues As Object, ByVal p_InitialValues As Object, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Dim changed As Boolean
    p_Error = ""
    changed = False
    If p_InitialValues Is Nothing Then
        changed = True
    Else
        changed = (DictValue(p_CurrentValues, "DESCRIPCION") <> DictValue(p_InitialValues, "DESCRIPCION"))
        If Not changed Then changed = (DictValue(p_CurrentValues, "NModificado") <> DictValue(p_InitialValues, "NModificado"))
        If Not changed Then changed = (DictValue(p_CurrentValues, "FechaFirmaModificado") <> DictValue(p_InitialValues, "FechaFirmaModificado"))
        If Not changed Then changed = (DictValue(p_CurrentValues, "FechaFinModificado") <> DictValue(p_InitialValues, "FechaFinModificado"))
    End If
    Set payload = CreateObject("Scripting.Dictionary")
    payload("hasChanges") = changed
    Modificado_HasChanges = Envelope(payload)
End Function

Public Function Modificado_RegisterDecision(ByVal p_HasChanges As Boolean, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("canRegister") = p_HasChanges
    If Not p_HasChanges Then p_Error = "No hay cambios que guardar"
    Modificado_RegisterDecision = BuildJsonPayload(p_HasChanges, payload, p_Error, logs)
End Function
