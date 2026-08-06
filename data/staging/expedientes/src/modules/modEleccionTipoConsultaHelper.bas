Attribute VB_Name = "modEleccionTipoConsultaHelper"
Option Compare Database
Option Explicit

' Pure-data helpers for Form_FormEleccionTipoConsulta.
' Public prefix: EleccionTipoConsulta_*

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

Public Function EleccionTipoConsulta_ParseOpenArgs(ByVal p_OpenArgs As String, ByVal p_DefaultList As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""

    If Len(Trim$(p_OpenArgs)) = 0 Then
        payload("nombreLista") = p_DefaultList
        payload("nombreFormulario") = "FormExpedientesGestion"
        EleccionTipoConsulta_ParseOpenArgs = Envelope(payload)
        Exit Function
    End If

    Dim parts As Variant
    parts = Split(p_OpenArgs, "|")
    If UBound(parts) <> 1 Then
        p_Error = "El formulario se ha abierto con parámetros insuficientes"
        EleccionTipoConsulta_ParseOpenArgs = Envelope(Nothing, p_Error)
        Exit Function
    End If

    payload("nombreLista") = CStr(parts(0))
    payload("nombreFormulario") = CStr(parts(1))
    EleccionTipoConsulta_ParseOpenArgs = Envelope(payload)
End Function

Public Function EleccionTipoConsulta_InitialListState(Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    payload("rowSource") = "Consulta Completa;Consulta Simple"
    payload("defaultIndex") = 1
    payload("executeEnabled") = False
    payload("updateEnabled") = False
    EleccionTipoConsulta_InitialListState = Envelope(payload)
End Function

Public Function EleccionTipoConsulta_SelectType(ByVal p_TipoConsulta As String, ByVal p_CollectionsLoaded As Boolean, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""

    Select Case Trim$(p_TipoConsulta)
        Case "Consulta Completa"
            payload("fieldSet") = "completo"
        Case "Consulta Simple"
            payload("fieldSet") = "simple"
        Case ""
            payload("fieldSet") = "none"
            payload("executeEnabled") = False
            payload("updateEnabled") = False
            payload("loadCollections") = False
            EleccionTipoConsulta_SelectType = Envelope(payload)
            Exit Function
        Case Else
            p_Error = "Tipo no codificado"
            EleccionTipoConsulta_SelectType = Envelope(Nothing, p_Error)
            Exit Function
    End Select

    payload("executeEnabled") = True
    payload("updateEnabled") = True
    payload("loadCollections") = Not p_CollectionsLoaded
    EleccionTipoConsulta_SelectType = Envelope(payload)
End Function

Public Function EleccionTipoConsulta_ListLoadPlan(ByVal p_NombreLista As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    Select Case Trim$(p_NombreLista)
        Case "Todos"
            payload("loadAM") = True
            payload("loadLotes") = True
            payload("loadBasados") = True
            payload("loadTecnica") = False
        Case "ListaAMEIndividual"
            payload("loadAM") = True
            payload("loadLotes") = False
            payload("loadBasados") = False
            payload("loadTecnica") = False
        Case "ListaLotes"
            payload("loadAM") = False
            payload("loadLotes") = True
            payload("loadBasados") = False
            payload("loadTecnica") = False
        Case "ListaBasados"
            payload("loadAM") = False
            payload("loadLotes") = False
            payload("loadBasados") = True
            payload("loadTecnica") = False
        Case "ListaFiltrados"
            payload("loadAM") = False
            payload("loadLotes") = False
            payload("loadBasados") = False
            payload("loadTecnica") = True
        Case Else
            p_Error = "Lista no reconocida"
            EleccionTipoConsulta_ListLoadPlan = Envelope(Nothing, p_Error)
            Exit Function
    End Select
    EleccionTipoConsulta_ListLoadPlan = Envelope(payload)
End Function

Public Function EleccionTipoConsulta_ExportRequest(ByVal p_NombreListaParaInforme As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    payload("incluirDerivados") = (p_NombreListaParaInforme = "Todos")
    EleccionTipoConsulta_ExportRequest = Envelope(payload)
End Function
