Attribute VB_Name = "mod0BDGestorEntidadesHelper"
Option Compare Database
Option Explicit

' Pure-data helpers for Form_Form0BDGestorEntidades.
' Public prefix: GestorEntidades_*

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

Public Function GestorEntidades_ListRowSource(ByVal p_Opciones As Object, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""

    Dim rowSource As String
    Dim countRows As Long
    If Not p_Opciones Is Nothing Then
        Dim key As Variant
        For Each key In p_Opciones.Keys
            If Len(rowSource) > 0 Then rowSource = rowSource & vbCrLf
            rowSource = rowSource & CStr(key) & ";" & Replace(CStr(p_Opciones(key)), ";", ":")
            countRows = countRows + 1
        Next key
    End If

    payload("rowSource") = rowSource
    payload("count") = countRows
    GestorEntidades_ListRowSource = Envelope(payload)
End Function

Public Function GestorEntidades_SelectionState(ByVal p_SelectedID As String, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    payload("hasSelection") = (Len(Trim$(p_SelectedID)) > 0)
    payload("enableOpen") = CBool(payload("hasSelection"))
    GestorEntidades_SelectionState = Envelope(payload)
End Function

Public Function GestorEntidades_OpenDecision(ByVal p_SelectedID As String, ByVal p_FormMap As Object, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""

    Dim selectedId As String
    selectedId = Trim$(p_SelectedID)
    If Len(selectedId) = 0 Then
        p_Error = "No se ha seleccionado ninguna opción"
        GestorEntidades_OpenDecision = Envelope(Nothing, p_Error)
        Exit Function
    End If

    Dim formName As String
    If Not p_FormMap Is Nothing Then
        If p_FormMap.Exists(selectedId) Then formName = CStr(p_FormMap(selectedId))
    End If
    If Len(Trim$(formName)) = 0 Then
        p_Error = "No se ha seleccionado ninguna entidad reconocida"
        GestorEntidades_OpenDecision = Envelope(Nothing, p_Error)
        Exit Function
    End If

    payload("formName") = formName
    payload("shouldOpen") = True
    GestorEntidades_OpenDecision = Envelope(payload)
End Function

Public Function GestorEntidades_DoubleClickDecision(ByVal p_OpenEnabled As Boolean, Optional ByRef p_Error As String) As String
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    p_Error = ""
    If p_OpenEnabled Then payload("action") = "open" Else payload("action") = "none"
    GestorEntidades_DoubleClickDecision = Envelope(payload)
End Function
