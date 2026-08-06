Attribute VB_Name = "modExpedienteSuministradoresHelper"
Option Compare Database
Option Explicit

' Pure-data helpers for Form_FormExpedienteSuministradores.
' Public prefix: ExpedienteSuministradores_*

Private Function BuildJsonPayload(ByVal p_Ok As Boolean, ByVal p_Payload As Object, ByVal p_ErrorMsg As String, ByRef p_Logs() As String) As String
    Dim payloadJson As String
    If p_Payload Is Nothing Then payloadJson = "null" Else payloadJson = JsonConverter.ConvertToJson(p_Payload)
    Dim errorJson As String
    If p_Ok Then errorJson = "null" Else errorJson = """" & TestHelper.EscapeJsonString(p_ErrorMsg) & """"
    BuildJsonPayload = "{""ok"":" & LCase$(CStr(p_Ok)) & ",""value"":null,""payload"":" & payloadJson & ",""error"":" & errorJson & ",""logs"":" & TestHelper.JsonStringArray(p_Logs) & "}"
End Function

Public Function ExpedienteSuministradores_ExtraerTag(ByVal p_TagText As String, ByVal p_KeyName As String, Optional ByRef p_Error As String) As String
    On Error GoTo errores
    p_Error = ""
    Dim parts() As String
    parts = Split("" & p_TagText, ";")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim kv() As String
        kv = Split(parts(i), "=")
        If UBound(kv) >= 1 Then
            If UCase$(Trim$(kv(0))) = UCase$(Trim$(p_KeyName)) Then
                ExpedienteSuministradores_ExtraerTag = Trim$(kv(1))
                Exit Function
            End If
        End If
    Next
    ExpedienteSuministradores_ExtraerTag = ""
    Exit Function
errores:
    p_Error = "ExtraerTag: " & Err.Description
    ExpedienteSuministradores_ExtraerTag = ""
End Function

Public Function ExpedienteSuministradores_ObtenerTextoSocios(ByVal p_ColNombres As Object, ByVal p_NombreActual As String, Optional ByRef p_Error As String) As String
    On Error GoTo errores
    p_Error = ""
    If p_ColNombres Is Nothing Then Exit Function
    If p_ColNombres.Count < 2 Then Exit Function
    Dim vItem As Variant
    Dim resultado As String
    For Each vItem In p_ColNombres
        If CStr(vItem) <> p_NombreActual Then resultado = resultado & CStr(vItem) & ", "
    Next
    If Len(resultado) > 0 Then
        resultado = Left$(resultado, Len(resultado) - 2)
        ExpedienteSuministradores_ObtenerTextoSocios = " (UTE con: " & resultado & ")"
    End If
    Exit Function
errores:
    p_Error = "ObtenerTextoSocios: " & Err.Description
    ExpedienteSuministradores_ObtenerTextoSocios = ""
End Function

Public Function ExpedienteSuministradores_GestionarBotonesState(ByVal p_HaySeleccion As Boolean, ByVal p_Key As String, ByVal p_AllowEdits As Boolean, Optional ByRef p_Error As String) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""
    Dim esRaiz As Boolean
    esRaiz = False
    If p_HaySeleccion Then esRaiz = (Left$("" & p_Key, 4) = "ROOT")
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("eliminarEnabled") = (p_HaySeleccion And Not esRaiz And p_AllowEdits)
    payload("verDetalleEnabled") = (p_HaySeleccion And Not esRaiz)
    ExpedienteSuministradores_GestionarBotonesState = BuildJsonPayload(True, payload, "", logs)
End Function

Public Function ExpedienteSuministradores_ResolveDropTarget(ByVal p_TargetKey As String, ByVal p_TargetTag As String, ByVal p_IsEmptyTarget As Boolean, Optional ByRef p_Error As String) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    If p_IsEmptyTarget Or p_TargetKey = "ROOT_ORGANO" Then
        payload("tipo") = "ROOT_CONTR"
        payload("parentRelID") = Null
        payload("keyDest") = "ROOT_ORGANO"
    Else
        payload("tipo") = "CHILD"
        payload("parentRelID") = ExpedienteSuministradores_ExtraerTag(p_TargetTag, "RELID", p_Error)
        payload("keyDest") = p_TargetKey
    End If
    ExpedienteSuministradores_ResolveDropTarget = BuildJsonPayload(True, payload, "", logs)
End Function
