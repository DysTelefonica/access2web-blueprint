Attribute VB_Name = "modTareasHelper"
Option Compare Database
Option Explicit

' Pure-data helpers for Form_FormTareas.
' Public prefix: Tareas_*

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

Public Function Tareas_ListHeader(Optional ByRef p_Error As String) As String
    Dim payload As Object
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("rowSource") = "IDExpediente;Estado;Tipo;Nemotécnico;NºExp;Título"
    payload("emptyCaption") = "Registros filtrados (0)"
    Tareas_ListHeader = Envelope(payload)
End Function

Public Function Tareas_TaskTypeIsValid(ByVal p_TipoTarea As Variant, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("isValid") = IsNumeric(p_TipoTarea)
    If Not CBool(payload("isValid")) Then p_Error = "Seleccione un tipo de tarea"
    Tareas_TaskTypeIsValid = BuildJsonPayload(CBool(payload("isValid")), payload, p_Error, logs)
End Function

Public Function Tareas_ReportAction(ByVal p_TipoTarea As Long, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    Select Case p_TipoTarea
        Case EnumTipoTarea.EstadoDesconocido
            payload("action") = "EstadoDesconocido"
        Case EnumTipoTarea.APuntoDeRecepcionarCompleto
            payload("action") = "APuntoDeRecepcionarCompleto"
        Case EnumTipoTarea.APuntoDeRecepcionarHito
            payload("action") = "APuntoDeRecepcionarHito"
        Case EnumTipoTarea.AdjudicadoSinContrato
            payload("action") = "AdjudicadoSinContrato"
        Case EnumTipoTarea.AdjudicadosTSOLSinCodS4H
            payload("action") = "AdjudicadosTSOLSinCodS4H"
        Case EnumTipoTarea.EnFaseOfertaPorMuchoTiempo
            payload("action") = "EnFaseOfertaPorMuchoTiempo"
        Case Else
            p_Error = "Tipo de tarea no codificada"
            payload("action") = ""
    End Select
    Tareas_ReportAction = BuildJsonPayload(Len(p_Error) = 0, payload, p_Error, logs)
End Function

Public Function Tareas_SelectionState(ByVal p_SelectedID As String, ByVal p_CurrentID As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    p_Error = ""
    Set payload = CreateObject("Scripting.Dictionary")
    payload("hasSelection") = (Len(Trim$(p_SelectedID)) > 0)
    payload("requiresLoad") = (Len(Trim$(p_SelectedID)) > 0 And Trim$(p_SelectedID) <> Trim$(p_CurrentID))
    payload("enableDetail") = CBool(payload("hasSelection"))
    payload("enableInforme") = CBool(payload("hasSelection"))
    Tareas_SelectionState = Envelope(payload)
End Function
