Attribute VB_Name = "Test_ExpedienteHitosHelper"
Option Compare Database
Option Explicit

' Test_ExpedienteHitosHelper — canonical JSON TDD atoms for PR #40.

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
    Set dto("ColHitos") = Nothing
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

Public Function Test_ExpedienteHitosHelper_AltaHito_SadEmptyFecha() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim json As String
    json = modExpedienteHitosHelper.ExpedienteHitos_AltaHito(BuildDTOWithExpediente(), "", "desc", "", "", p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false"
    Test_ExpedienteHitosHelper_AltaHito_SadEmptyFecha = BuildOk("empty-fecha-rejected", logs)
    Exit Function
EH:
    Test_ExpedienteHitosHelper_AltaHito_SadEmptyFecha = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteHitosHelper_EliminarHito_SadNotAdmin() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim json As String
    json = modExpedienteHitosHelper.ExpedienteHitos_EliminarHito(BuildDTOWithExpediente(), "2026-01-01", False, p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false"
    If p_Error <> "Operacion no autorizada" Then Err.Raise 1002, , "expected unauthorized error"
    Test_ExpedienteHitosHelper_EliminarHito_SadNotAdmin = BuildOk("not-admin-rejected", logs)
    Exit Function
EH:
    Test_ExpedienteHitosHelper_EliminarHito_SadNotAdmin = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteHitosHelper_RellenarListaHitos_EdgeEmptyCol() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteHitosHelper.ExpedienteHitos_RellenarListaHitos(BuildEmptyDTO(), p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CLng(payload("count")) <> 0 Then Err.Raise 1001, , "expected count=0"
    If CStr(payload("rowSource")) <> "Fecha;Descripcion;Fecha G.;Importe" Then Err.Raise 1002, , "expected hitos header"
    Test_ExpedienteHitosHelper_RellenarListaHitos_EdgeEmptyCol = BuildOk("empty-hitos", logs)
    Exit Function
EH:
    Test_ExpedienteHitosHelper_RellenarListaHitos_EdgeEmptyCol = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteHitosHelper_RellenarListaHitos_EdgeMissingKey() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH
    Dim dto As Object
    Set dto = CreateObject("Scripting.Dictionary")
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteHitosHelper.ExpedienteHitos_RellenarListaHitos(dto, p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CLng(payload("count")) <> 0 Then Err.Raise 1001, , "expected count=0"
    If CStr(payload("rowSource")) <> "Fecha;Descripcion;Fecha G.;Importe" Then Err.Raise 1002, , "expected hitos header"
    Test_ExpedienteHitosHelper_RellenarListaHitos_EdgeMissingKey = BuildOk("missing-key-safe-default", logs)
    Exit Function
EH:
    Test_ExpedienteHitosHelper_RellenarListaHitos_EdgeMissingKey = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteHitosHelper_RellenarListaHitos_EdgeScalarSlot() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH
    Dim dto As Object
    Set dto = CreateObject("Scripting.Dictionary")
    dto("ColHitos") = "not-an-object"
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteHitosHelper.ExpedienteHitos_RellenarListaHitos(dto, p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CLng(payload("count")) <> 0 Then Err.Raise 1001, , "expected count=0"
    If CStr(payload("rowSource")) <> "Fecha;Descripcion;Fecha G.;Importe" Then Err.Raise 1002, , "expected hitos header"
    Test_ExpedienteHitosHelper_RellenarListaHitos_EdgeScalarSlot = BuildOk("scalar-slot-safe-default", logs)
    Exit Function
EH:
    Test_ExpedienteHitosHelper_RellenarListaHitos_EdgeScalarSlot = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteHitosHelper_RellenarListaHitos_EdgeNonDictionaryDTO() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH
    Dim dto As Object
    Set dto = New Collection
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteHitosHelper.ExpedienteHitos_RellenarListaHitos(dto, p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CLng(payload("count")) <> 0 Then Err.Raise 1001, , "expected count=0"
    If CStr(payload("rowSource")) <> "Fecha;Descripcion;Fecha G.;Importe" Then Err.Raise 1002, , "expected hitos header"
    Test_ExpedienteHitosHelper_RellenarListaHitos_EdgeNonDictionaryDTO = BuildOk("non-dictionary-safe-default", logs)
    Exit Function
EH:
    Test_ExpedienteHitosHelper_RellenarListaHitos_EdgeNonDictionaryDTO = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteHitosHelper_RellenarListaHitos_SadUnexpectedDTOError() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH
    Dim dto As TestUnexpectedErrorDTO
    Set dto = New TestUnexpectedErrorDTO
    Dim p_Error As String
    Dim json As String
    json = modExpedienteHitosHelper.ExpedienteHitos_RellenarListaHitos(dto, p_Error)
    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed Is Nothing Then Err.Raise 1001, , "expected structured JSON object"
    If Not parsed.Exists("ok") Then Err.Raise 1002, , "expected ok field"
    If VarType(parsed("ok")) <> vbBoolean Then Err.Raise 1003, , "expected Boolean ok field"
    If CBool(parsed("ok")) <> False Then Err.Raise 1004, , "expected ok=false"
    Dim expectedError As String
    expectedError = "RellenarListaHitos: controlled unexpected DTO Exists failure"
    If p_Error <> expectedError Then Err.Raise 1005, , "unexpected public error"
    If Not parsed.Exists("error") Then Err.Raise 1006, , "expected error field"
    If IsNull(parsed("error")) Then Err.Raise 1007, , "expected non-null error field"
    If CStr(parsed("error")) <> expectedError Then Err.Raise 1008, , "unexpected JSON error"
    Test_ExpedienteHitosHelper_RellenarListaHitos_SadUnexpectedDTOError = BuildOk("unexpected-error-propagated", logs)
    Exit Function
EH:
    Test_ExpedienteHitosHelper_RellenarListaHitos_SadUnexpectedDTOError = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteHitosHelper_EstablecerDatos_HappyAdminEdit() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteHitosHelper.ExpedienteHitos_EstablecerDatos(BuildDTOWithExpediente(), True, True, p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CBool(payload("expedienteOK")) <> True Then Err.Raise 1001, , "expected expedienteOK=true"
    If CBool(payload("ejecutivosEnabled")) <> True Then Err.Raise 1002, , "expected ejecutivosEnabled=true"
    Test_ExpedienteHitosHelper_EstablecerDatos_HappyAdminEdit = BuildOk("admin-edit", logs)
    Exit Function
EH:
    Test_ExpedienteHitosHelper_EstablecerDatos_HappyAdminEdit = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteHitosHelper_Form_Load_SadDTOWithoutExpediente() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH
    Dim p_Error As String
    Dim json As String
    json = modExpedienteHitosHelper.ExpedienteHitos_Form_Load(BuildEmptyDTO(), True, True, p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false"
    Test_ExpedienteHitosHelper_Form_Load_SadDTOWithoutExpediente = BuildOk("missing-expediente-rejected", logs)
    Exit Function
EH:
    Test_ExpedienteHitosHelper_Form_Load_SadDTOWithoutExpediente = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteHitosHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH
    logs(0) = Test_ExpedienteHitosHelper_AltaHito_SadEmptyFecha()
    logs(1) = Test_ExpedienteHitosHelper_EliminarHito_SadNotAdmin()
    logs(2) = Test_ExpedienteHitosHelper_RellenarListaHitos_EdgeEmptyCol()
    logs(3) = Test_ExpedienteHitosHelper_EstablecerDatos_HappyAdminEdit()
    logs(4) = Test_ExpedienteHitosHelper_Form_Load_SadDTOWithoutExpediente()
    Test_ExpedienteHitosHelper_RunAll = BuildOk("5 atoms executed", logs)
    Exit Function
EH:
    Test_ExpedienteHitosHelper_RunAll = BuildFail(Err.Description, logs)
End Function
