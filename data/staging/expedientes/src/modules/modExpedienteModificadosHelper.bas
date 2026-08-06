Attribute VB_Name = "modExpedienteModificadosHelper"
Option Compare Database
Option Explicit

' Pure-data helpers for Form_FormExpedienteModificados.
' Public prefix: ExpedienteModificados_*

Private Const EXPEDIENTEMOD_FIELD_SEP As String = ";"
Private Const EXPEDIENTEMOD_HEADERS As String = "ID;Nº;Descripción;Firma;F.Fin"
Private Const VBA_ERROR_OBJECT_REQUIRED As Long = 424
Private Const VBA_ERROR_MEMBER_NOT_SUPPORTED As Long = 438

Private Function BuildJsonPayload(ByVal p_Ok As Boolean, ByVal p_Payload As Object, ByVal p_ErrorMsg As String, ByRef p_Logs() As String) As String
    Dim payloadJson As String
    If p_Payload Is Nothing Then
        payloadJson = "null"
    Else
        payloadJson = JsonConverter.ConvertToJson(p_Payload)
    End If
    Dim errorJson As String
    If p_Ok Then
        errorJson = "null"
    Else
        errorJson = """" & TestHelper.EscapeJsonString(p_ErrorMsg) & """"
    End If
    Dim okJson As String
    If p_Ok Then
        okJson = "true"
    Else
        okJson = "false"
    End If
    BuildJsonPayload = "{""ok"":" & okJson & ",""value"":null,""payload"":" & payloadJson & ",""error"":" & errorJson & ",""logs"":" & TestHelper.JsonStringArray(p_Logs) & "}"
End Function

Private Function GetDTOSafely(ByVal p_DTO As Object, ByVal p_Key As String) As Object
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim errHelpFile As String
    Dim errHelpContext As Long

    On Error GoTo HandleError
    If p_DTO Is Nothing Then Exit Function
    If Not p_DTO.Exists(p_Key) Then Exit Function
    If Not IsObject(p_DTO(p_Key)) Then Exit Function
    Set GetDTOSafely = p_DTO(p_Key)
    Exit Function

HandleError:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    errHelpFile = Err.HelpFile
    errHelpContext = Err.HelpContext

    If errNumber = VBA_ERROR_OBJECT_REQUIRED Then Exit Function
    If errNumber = VBA_ERROR_MEMBER_NOT_SUPPORTED Then Exit Function

    Err.Raise errNumber, errSource, errDescription, errHelpFile, errHelpContext
End Function

Private Function SafeRowText(ByVal p_Value As Variant) As String
    SafeRowText = Replace("" & Nz(p_Value, ""), EXPEDIENTEMOD_FIELD_SEP, ":")
End Function

Public Function ExpedienteModificados_FormLoadState(ByVal p_AbiertoParaEditar As Boolean, ByVal p_EsAdministrador As Boolean, Optional ByRef p_Error As String) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    p_Error = ""
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("allowEdits") = (p_AbiertoParaEditar And p_EsAdministrador)
    payload("parentRegistrarVisible") = False
    ExpedienteModificados_FormLoadState = BuildJsonPayload(True, payload, "", logs)
End Function

Public Function ExpedienteModificados_RellenarListaModificados(ByVal p_DTO As Object, Optional ByRef p_Error As String) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    p_Error = ""
    On Error GoTo errores
    Dim col As Object
    Set col = GetDTOSafely(p_DTO, "ColModificados")
    Dim rowSource As String
    rowSource = EXPEDIENTEMOD_HEADERS
    If Not col Is Nothing Then
        Dim key As Variant
        For Each key In col
            Dim item As Object
            Set item = col(key)
            rowSource = rowSource & vbCrLf & SafeRowText(item.IDExpedienteModificado) & EXPEDIENTEMOD_FIELD_SEP & SafeRowText(item.NModificado) & EXPEDIENTEMOD_FIELD_SEP & SafeRowText(item.DESCRIPCION) & EXPEDIENTEMOD_FIELD_SEP & SafeRowText(item.FechaFirmaModificado) & EXPEDIENTEMOD_FIELD_SEP & SafeRowText(item.FechaFinModificado)
        Next
    End If
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("rowSource") = rowSource
    ExpedienteModificados_RellenarListaModificados = BuildJsonPayload(True, payload, "", logs)
    Exit Function
errores:
    p_Error = "RellenarListaModificados: " & Err.Description
    ExpedienteModificados_RellenarListaModificados = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

Public Function ExpedienteModificados_RegistrarModificado(ByVal p_DTO As Object, ByVal p_Descripcion As String, ByVal p_FechaFinModificado As Variant, ByVal p_FechaFirmaModificado As Variant, ByVal p_NModificado As Variant, Optional ByRef p_Error As String) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    p_Error = ""
    On Error GoTo errores
    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente")
    If expObj Is Nothing Then
        p_Error = "RegistrarModificado: p_DTO.Expediente is Nothing"
        ExpedienteModificados_RegistrarModificado = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    Dim modificado As New ExpedienteModificado
    modificado.DESCRIPCION = "" & p_Descripcion
    modificado.FechaFinModificado = "" & p_FechaFinModificado
    modificado.FechaFirmaModificado = "" & p_FechaFirmaModificado
    modificado.NModificado = "" & p_NModificado
    Dim expOp As New ExpedienteOperaciones
    Set expOp.Expediente = expObj
    Dim estado As String
    estado = expOp.RegistrarModificado(p_Modificado:=modificado, p_Error:=p_Error)
    If p_Error <> "" Then ExpedienteModificados_RegistrarModificado = BuildJsonPayload(False, Nothing, p_Error, logs): Exit Function
    On Error Resume Next
    Set expObj.Modificados = Nothing
    Dim colNew As Object
    Set colNew = expObj.Modificados
    On Error GoTo errores
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("registrado") = (estado = "1")
    payload("colRefreshed") = False
    If estado = "1" Then
        If Not colNew Is Nothing Then
            Set p_DTO("ColModificados") = colNew
            payload("colRefreshed") = True
        End If
    End If
    ExpedienteModificados_RegistrarModificado = BuildJsonPayload(True, payload, "", logs)
    Exit Function
errores:
    If p_Error = "" Then p_Error = "RegistrarModificado: " & Err.Description
    ExpedienteModificados_RegistrarModificado = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

Public Function ExpedienteModificados_EliminarModificado(ByVal p_DTO As Object, ByVal p_IDExpedienteModificado As String, ByVal p_EsAdministrador As Boolean, Optional ByRef p_Error As String) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""
    On Error GoTo errores
    If Not p_EsAdministrador Then
        p_Error = "Operacion no autorizada"
        ExpedienteModificados_EliminarModificado = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    Dim idMod As String
    idMod = Trim$("" & p_IDExpedienteModificado)
    If Len(idMod) = 0 Then
        Dim payloadNoop As Object
        Set payloadNoop = CreateObject("Scripting.Dictionary")
        payloadNoop("eliminado") = False
        ExpedienteModificados_EliminarModificado = BuildJsonPayload(True, payloadNoop, "", logs)
        Exit Function
    End If
    Dim expObj As Object
    Set expObj = GetDTOSafely(p_DTO, "Expediente")
    If expObj Is Nothing Then
        p_Error = "EliminarModificado: p_DTO.Expediente is Nothing"
        ExpedienteModificados_EliminarModificado = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If
    Dim expOp As New ExpedienteOperaciones
    Set expOp.Expediente = expObj
    expOp.EliminarModificado p_IDExpedienteModificado:=idMod, p_Error:=p_Error
    If p_Error <> "" Then ExpedienteModificados_EliminarModificado = BuildJsonPayload(False, Nothing, p_Error, logs): Exit Function
    On Error Resume Next
    Set expObj.Modificados = Nothing
    Dim colNew As Object
    Set colNew = expObj.Modificados
    On Error GoTo errores
    If Not colNew Is Nothing Then Set p_DTO("ColModificados") = colNew
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")
    payload("eliminado") = True
    ExpedienteModificados_EliminarModificado = BuildJsonPayload(True, payload, "", logs)
    Exit Function
errores:
    If p_Error = "" Then p_Error = "EliminarModificado: " & Err.Description
    ExpedienteModificados_EliminarModificado = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function
