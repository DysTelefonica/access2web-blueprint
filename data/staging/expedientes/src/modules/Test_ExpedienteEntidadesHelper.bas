Attribute VB_Name = "Test_ExpedienteEntidadesHelper"
Option Compare Database
Option Explicit

' Test_ExpedienteEntidadesHelper — canonical JSON TDD atoms for PR #40.
' Pure-data tests only: no UI dependencies or form automation.

Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

Private Function BuildEmptyDTO() As Object
    Dim dto As Object
    Set dto = CreateObject("Scripting.Dictionary")
    dto("Expediente") = Nothing
    dto("ColComerciales") = Nothing
    dto("ColCPVs") = Nothing
    dto("ColLugaresEjecucion") = Nothing
    dto("ColPECALES") = Nothing
    dto("ColRACs") = Nothing
    dto("ColResponsables") = Nothing
    dto("ColAnualidades") = Nothing
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

Private Function ErrorOf(ByVal p_Json As String) As String
    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(p_Json)
    ErrorOf = CStr(parsed("error"))
End Function

Public Function Test_ExpedienteEntidadesHelper_RellenarListaComerciales_EdgeEmptyCollection() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH
    Dim p_Error As String
    Dim json As String
    json = modExpedienteEntidadesHelper.ExpedienteEntidades_RellenarListaComerciales(BuildEmptyDTO(), p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    Dim payload As Object
    Set payload = PayloadOf(json)
    If CLng(payload("count")) <> 0 Then Err.Raise 1001, , "expected count=0"
    If CStr(payload("rowSource")) <> "" Then Err.Raise 1002, , "expected empty rowSource"
    Test_ExpedienteEntidadesHelper_RellenarListaComerciales_EdgeEmptyCollection = BuildOk("empty-comerciales", logs)
    Exit Function
EH:
    Test_ExpedienteEntidadesHelper_RellenarListaComerciales_EdgeEmptyCollection = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteEntidadesHelper_RellenarListaResponsables_EdgeEmptyCollection() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH
    Dim p_Error As String
    Dim json As String
    json = modExpedienteEntidadesHelper.ExpedienteEntidades_RellenarListaResponsables(BuildEmptyDTO(), p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    Dim payload As Object
    Set payload = PayloadOf(json)
    If CLng(payload("count")) <> 0 Then Err.Raise 1001, , "expected count=0"
    If CStr(payload("rowSource")) <> "ID;Nombre;JP;Aviso" Then Err.Raise 1002, , "expected responsables header"
    Test_ExpedienteEntidadesHelper_RellenarListaResponsables_EdgeEmptyCollection = BuildOk("empty-responsables", logs)
    Exit Function
EH:
    Test_ExpedienteEntidadesHelper_RellenarListaResponsables_EdgeEmptyCollection = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteEntidadesHelper_EstablecerDatos_HappyAdminEdit() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteEntidadesHelper.ExpedienteEntidades_EstablecerDatos(BuildDTOWithExpediente(), True, True, p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CBool(payload("expedienteOK")) <> True Then Err.Raise 1001, , "expected expedienteOK=true"
    If CBool(payload("ejecutivosEnabled")) <> True Then Err.Raise 1002, , "expected ejecutivosEnabled=true"
    Test_ExpedienteEntidadesHelper_EstablecerDatos_HappyAdminEdit = BuildOk("admin-edit", logs)
    Exit Function
EH:
    Test_ExpedienteEntidadesHelper_EstablecerDatos_HappyAdminEdit = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteEntidadesHelper_Form_Load_SadDTOWithoutExpediente() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH
    Dim p_Error As String
    Dim json As String
    json = modExpedienteEntidadesHelper.ExpedienteEntidades_Form_Load(BuildEmptyDTO(), True, True, p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false"
    If Len(ErrorOf(json)) = 0 Then Err.Raise 1002, , "expected non-empty error"
    Test_ExpedienteEntidadesHelper_Form_Load_SadDTOWithoutExpediente = BuildOk("missing-expediente-rejected", logs)
    Exit Function
EH:
    Test_ExpedienteEntidadesHelper_Form_Load_SadDTOWithoutExpediente = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteEntidadesHelper_CambiarLineaListaResponsables_HappyLineFound() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH
    Dim p_Error As String
    Dim payload As Object
    Set payload = PayloadOf(modExpedienteEntidadesHelper.ExpedienteEntidades_CambiarLineaListaResponsables("A" & vbCrLf & "B", "B", "C", p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CBool(payload("replaced")) <> True Then Err.Raise 1001, , "expected replaced=true"
    If InStr(CStr(payload("rowSource")), "C") = 0 Then Err.Raise 1002, , "expected replacement line"
    Test_ExpedienteEntidadesHelper_CambiarLineaListaResponsables_HappyLineFound = BuildOk("line-replaced", logs)
    Exit Function
EH:
    Test_ExpedienteEntidadesHelper_CambiarLineaListaResponsables_HappyLineFound = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteEntidadesHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH
    logs(0) = Test_ExpedienteEntidadesHelper_RellenarListaComerciales_EdgeEmptyCollection()
    logs(1) = Test_ExpedienteEntidadesHelper_RellenarListaResponsables_EdgeEmptyCollection()
    logs(2) = Test_ExpedienteEntidadesHelper_EstablecerDatos_HappyAdminEdit()
    logs(3) = Test_ExpedienteEntidadesHelper_Form_Load_SadDTOWithoutExpediente()
    logs(4) = Test_ExpedienteEntidadesHelper_CambiarLineaListaResponsables_HappyLineFound()
    Test_ExpedienteEntidadesHelper_RunAll = BuildOk("5 atoms executed", logs)
    Exit Function
EH:
    Test_ExpedienteEntidadesHelper_RunAll = BuildFail(Err.Description, logs)
End Function
