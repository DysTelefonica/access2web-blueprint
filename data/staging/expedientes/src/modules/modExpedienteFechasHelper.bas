Attribute VB_Name = "modExpedienteFechasHelper"
Option Compare Database
Option Explicit

' Pure-data helpers for Form_FormExpedienteFechas.
' Migrates the existing Helper_ExpedienteFechasLogic contract to the Phase 3.3
' naming convention: ExpedienteFechas_*.

Private Function BuildJsonPayload(ByVal p_Ok As Boolean, ByVal p_Payload As Object, ByVal p_ErrorMsg As String, ByRef p_Logs() As String) As String
    Dim payloadJson As String
    If p_Payload Is Nothing Then payloadJson = "null" Else payloadJson = JsonConverter.ConvertToJson(p_Payload)
    Dim errorJson As String
    If p_Ok Then errorJson = "null" Else errorJson = """" & TestHelper.EscapeJsonString(p_ErrorMsg) & """"
    BuildJsonPayload = "{""ok"":" & LCase$(CStr(p_Ok)) & ",""value"":null,""payload"":" & payloadJson & ",""error"":" & errorJson & ",""logs"":" & TestHelper.JsonStringArray(p_Logs) & "}"
End Function

Public Function ExpedienteFechas_CalcularFechaFinGarantia( _
    ByVal p_GarantiaMeses As Variant, _
    ByVal p_FechaCertificacion As Variant, _
    ByVal p_FechaFinContrato As Variant, _
    ByRef p_FechaFinGarantia As Variant, _
    Optional ByRef p_Error As String _
) As String
    On Error GoTo errores
    p_Error = ""
    If Not IsNumeric(p_GarantiaMeses) Then
        p_FechaFinGarantia = Null
        ExpedienteFechas_CalcularFechaFinGarantia = ""
        Exit Function
    End If
    If Not IsDate(p_FechaFinContrato) Then
        p_FechaFinGarantia = Null
        ExpedienteFechas_CalcularFechaFinGarantia = ""
        Exit Function
    End If
    If IsDate(p_FechaCertificacion) Then
        p_FechaFinGarantia = DateAdd("m", CDbl(p_GarantiaMeses), CDate(p_FechaCertificacion))
    Else
        p_FechaFinGarantia = DateAdd("m", CDbl(p_GarantiaMeses), CDate(p_FechaFinContrato))
    End If
    ExpedienteFechas_CalcularFechaFinGarantia = ""
    Exit Function
errores:
    p_Error = "ExpedienteFechas_CalcularFechaFinGarantia: " & Err.Description
    ExpedienteFechas_CalcularFechaFinGarantia = "ERR"
End Function

Public Function ExpedienteFechas_FormLoadState( _
    ByVal p_AbiertoParaEditar As Boolean, _
    ByVal p_EsAdministrador As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""
    Dim allowEdits As Boolean
    allowEdits = (p_AbiertoParaEditar And p_EsAdministrador)
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("allowEdits") = allowEdits
    payload("parentRegistrarVisible") = True
    payload("parentRegistrarEnabled") = allowEdits
    payload("perdidaEnabled") = False
    ExpedienteFechas_FormLoadState = BuildJsonPayload(True, payload, "", logs)
End Function

Public Function ExpedienteFechas_UnloadDecision( _
    ByVal p_AllowEdits As Boolean, _
    ByVal p_SnapshotInicial As String, _
    ByVal p_SnapshotActual As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("guardar") = (p_AllowEdits And p_SnapshotInicial <> p_SnapshotActual)
    ExpedienteFechas_UnloadDecision = BuildJsonPayload(True, payload, "", logs)
End Function
