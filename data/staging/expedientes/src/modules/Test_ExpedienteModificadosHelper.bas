Attribute VB_Name = "Test_ExpedienteModificadosHelper"
Option Compare Database
Option Explicit

' Test_ExpedienteModificadosHelper — canonical JSON TDD atoms for PR #40.

Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

Private Function BuildEmptyDTO() As Object
    Dim dto As Object
    Set dto = CreateObject("Scripting.Dictionary")
    Set dto("Expediente") = Nothing
    Set dto("ColModificados") = Nothing
    Set BuildEmptyDTO = dto
End Function

Private Function BuildDTOWithExpediente() As Object
    Dim dto As Object
    Set dto = BuildEmptyDTO()
    Dim expStub As Object
    Set expStub = CreateObject("Scripting.Dictionary")
    expStub("IDExpediente") = "EXP-TEST-001"
    Set dto("Expediente") = expStub
    Set BuildDTOWithExpediente = dto
End Function

Private Function PayloadOf(ByVal p_Json As String) As Object
    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(p_Json)
    Set PayloadOf = parsed("payload")
End Function

Public Function Test_ExpedienteModificadosHelper_FormLoadState_HappyAdminEditable() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteModificadosHelper.ExpedienteModificados_FormLoadState(True, True, p_Error))
    If CBool(payload("allowEdits")) <> True Then Err.Raise 1001, , "expected allowEdits=true"
    If CBool(payload("parentRegistrarVisible")) <> False Then Err.Raise 1002, , "expected parentRegistrarVisible=false"
    Test_ExpedienteModificadosHelper_FormLoadState_HappyAdminEditable = BuildOk("admin-edit", logs)
    Exit Function
EH:
    Test_ExpedienteModificadosHelper_FormLoadState_HappyAdminEditable = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteModificadosHelper_RellenarLista_EdgeEmptyDTO() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteModificadosHelper.ExpedienteModificados_RellenarListaModificados(BuildEmptyDTO(), p_Error))
    If CStr(payload("rowSource")) <> "ID;Nº;Descripción;Firma;F.Fin" Then Err.Raise 1001, , "expected header rowSource"
    Test_ExpedienteModificadosHelper_RellenarLista_EdgeEmptyDTO = BuildOk("empty-list", logs)
    Exit Function
EH:
    Test_ExpedienteModificadosHelper_RellenarLista_EdgeEmptyDTO = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteModificadosHelper_RellenarLista_EdgeMissingKey() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH
    Dim dto As Object
    Set dto = CreateObject("Scripting.Dictionary")
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteModificadosHelper.ExpedienteModificados_RellenarListaModificados(dto, p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CStr(payload("rowSource")) <> "ID;Nº;Descripción;Firma;F.Fin" Then Err.Raise 1001, , "expected header rowSource"
    Test_ExpedienteModificadosHelper_RellenarLista_EdgeMissingKey = BuildOk("missing-key-safe-default", logs)
    Exit Function
EH:
    Test_ExpedienteModificadosHelper_RellenarLista_EdgeMissingKey = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteModificadosHelper_RellenarLista_EdgeScalarSlot() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH
    Dim dto As Object
    Set dto = CreateObject("Scripting.Dictionary")
    dto("ColModificados") = "not-an-object"
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteModificadosHelper.ExpedienteModificados_RellenarListaModificados(dto, p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CStr(payload("rowSource")) <> "ID;Nº;Descripción;Firma;F.Fin" Then Err.Raise 1001, , "expected header rowSource"
    Test_ExpedienteModificadosHelper_RellenarLista_EdgeScalarSlot = BuildOk("scalar-slot-safe-default", logs)
    Exit Function
EH:
    Test_ExpedienteModificadosHelper_RellenarLista_EdgeScalarSlot = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteModificadosHelper_RellenarLista_EdgeNonDictionaryDTO() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH
    Dim dto As Object
    Set dto = New Collection
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteModificadosHelper.ExpedienteModificados_RellenarListaModificados(dto, p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CStr(payload("rowSource")) <> "ID;Nº;Descripción;Firma;F.Fin" Then Err.Raise 1001, , "expected header rowSource"
    Test_ExpedienteModificadosHelper_RellenarLista_EdgeNonDictionaryDTO = BuildOk("non-dictionary-safe-default", logs)
    Exit Function
EH:
    Test_ExpedienteModificadosHelper_RellenarLista_EdgeNonDictionaryDTO = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteModificadosHelper_RellenarLista_SadUnexpectedDTOError() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH
    Dim dto As TestUnexpectedErrorDTO
    Set dto = New TestUnexpectedErrorDTO
    Dim p_Error As String
    Dim json As String
    json = modExpedienteModificadosHelper.ExpedienteModificados_RellenarListaModificados(dto, p_Error)
    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed Is Nothing Then Err.Raise 1001, , "expected structured JSON object"
    If Not parsed.Exists("ok") Then Err.Raise 1002, , "expected ok field"
    If VarType(parsed("ok")) <> vbBoolean Then Err.Raise 1003, , "expected Boolean ok field"
    If CBool(parsed("ok")) <> False Then Err.Raise 1004, , "expected ok=false"
    Dim expectedError As String
    expectedError = "RellenarListaModificados: controlled unexpected DTO Exists failure"
    If p_Error <> expectedError Then Err.Raise 1005, , "unexpected public error"
    If Not parsed.Exists("error") Then Err.Raise 1006, , "expected error field"
    If IsNull(parsed("error")) Then Err.Raise 1007, , "expected non-null error field"
    If CStr(parsed("error")) <> expectedError Then Err.Raise 1008, , "unexpected JSON error"
    Test_ExpedienteModificadosHelper_RellenarLista_SadUnexpectedDTOError = BuildOk("unexpected-error-propagated", logs)
    Exit Function
EH:
    Test_ExpedienteModificadosHelper_RellenarLista_SadUnexpectedDTOError = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteModificadosHelper_Registrar_SadDTOWithoutExpediente() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim json As String
    json = modExpedienteModificadosHelper.ExpedienteModificados_RegistrarModificado(BuildEmptyDTO(), "desc", "", "", "1", p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false"
    Test_ExpedienteModificadosHelper_Registrar_SadDTOWithoutExpediente = BuildOk("missing-expediente-rejected", logs)
    Exit Function
EH:
    Test_ExpedienteModificadosHelper_Registrar_SadDTOWithoutExpediente = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteModificadosHelper_Eliminar_SadNotAdmin() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim json As String
    json = modExpedienteModificadosHelper.ExpedienteModificados_EliminarModificado(BuildDTOWithExpediente(), "900001", False, p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false"
    If p_Error <> "Operacion no autorizada" Then Err.Raise 1002, , "expected unauthorized error"
    Test_ExpedienteModificadosHelper_Eliminar_SadNotAdmin = BuildOk("not-admin-rejected", logs)
    Exit Function
EH:
    Test_ExpedienteModificadosHelper_Eliminar_SadNotAdmin = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteModificadosHelper_Eliminar_EdgeEmptyID() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteModificadosHelper.ExpedienteModificados_EliminarModificado(BuildDTOWithExpediente(), "", True, p_Error))
    If CBool(payload("eliminado")) <> False Then Err.Raise 1001, , "expected eliminado=false"
    Test_ExpedienteModificadosHelper_Eliminar_EdgeEmptyID = BuildOk("empty-id-noop", logs)
    Exit Function
EH:
    Test_ExpedienteModificadosHelper_Eliminar_EdgeEmptyID = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteModificadosHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH
    logs(0) = Test_ExpedienteModificadosHelper_FormLoadState_HappyAdminEditable()
    logs(1) = Test_ExpedienteModificadosHelper_RellenarLista_EdgeEmptyDTO()
    logs(2) = Test_ExpedienteModificadosHelper_Registrar_SadDTOWithoutExpediente()
    logs(3) = Test_ExpedienteModificadosHelper_Eliminar_SadNotAdmin()
    logs(4) = Test_ExpedienteModificadosHelper_Eliminar_EdgeEmptyID()
    Test_ExpedienteModificadosHelper_RunAll = BuildOk("5 atoms executed", logs)
    Exit Function
EH:
    Test_ExpedienteModificadosHelper_RunAll = BuildFail(Err.Description, logs)
End Function
