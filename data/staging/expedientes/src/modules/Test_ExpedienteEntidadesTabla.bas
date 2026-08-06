Attribute VB_Name = "Test_ExpedienteEntidadesTabla"
Option Compare Database
Option Explicit

Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

Private Function ParsedJson(ByVal p_Json As String) As Object
    Set ParsedJson = JsonConverter.ParseJson(p_Json)
End Function

Private Function BuildEntity( _
    ByVal p_IDEntidad As Long, _
    ByVal p_Nombre As String, _
    Optional ByVal p_Tipo As String = "Comercial" _
) As Object
    Dim entity As Object
    Set entity = CreateObject("Scripting.Dictionary")
    entity("IDEntidad") = p_IDEntidad
    entity("Nombre") = p_Nombre
    entity("Tipo") = p_Tipo
    Set BuildEntity = entity
End Function

Private Function BuildEntityWithoutID() As Object
    Dim entity As Object
    Set entity = CreateObject("Scripting.Dictionary")
    entity("Nombre") = "Missing ID"
    Set BuildEntityWithoutID = entity
End Function

Private Function BuildEntityCollection() As Object
    Dim entities As Object
    Set entities = CreateObject("Scripting.Dictionary")
    entities.Add "10", BuildEntity(10, "Comercial A", "Comercial")
    entities.Add "20", BuildEntity(20, "CPV B", "CPV")
    Set BuildEntityCollection = entities
End Function

Private Function BuildDTOExpediente(ByVal p_IDExpediente As Long) As Object
    Dim exp As Object
    Set exp = CreateObject("Scripting.Dictionary")
    exp("IDExpediente") = p_IDExpediente
    Set BuildDTOExpediente = exp
End Function

Private Function BuildDTOCollection( _
    ByVal p_Key As String, _
    ByVal p_Item As Object _
) As Object
    Dim col As Object
    Set col = CreateObject("Scripting.Dictionary")
    col.Add p_Key, p_Item
    Set BuildDTOCollection = col
End Function

Private Function BuildTypedEntity( _
    ByVal p_IDProp As String, _
    ByVal p_ID As String, _
    ByVal p_NameProp As String, _
    ByVal p_Name As String _
) As Object
    Dim entity As Object
    Set entity = CreateObject("Scripting.Dictionary")
    entity(p_IDProp) = p_ID
    entity(p_NameProp) = p_Name
    Set BuildTypedEntity = entity
End Function

Private Function BuildEntidadForTablaPayload( _
    ByVal p_IDEntidad As String, _
    ByVal p_Nombre As String, _
    ByVal p_Tipo As String _
) As Object
    Dim item As Object
    Set item = CreateObject("Scripting.Dictionary")
    item("IDEntidad") = p_IDEntidad
    item("Nombre") = p_Nombre
    item("Tipo") = p_Tipo
    Set BuildEntidadForTablaPayload = item
End Function

Private Function SafeCStr(ByVal p_Value As Variant) As String
    If IsNull(p_Value) Or IsEmpty(p_Value) Then
        SafeCStr = ""
    Else
        SafeCStr = CStr(p_Value)
    End If
End Function

Private Function BuildEntidadesPayloadFromDTO(ByVal p_DTO As Object) As Object
    Dim entidades As Object
    Set entidades = CreateObject("Scripting.Dictionary")
    entidades.CompareMode = TextCompare

    Dim col As Object
    Dim item As Object
    Dim key As Variant

    If p_DTO Is Nothing Then
        Set BuildEntidadesPayloadFromDTO = entidades
        Exit Function
    End If

    If p_DTO.Exists("ColComerciales") Then
        Set col = p_DTO("ColComerciales")
        For Each key In col.Keys
            Set item = col(key)
            If Not item Is Nothing Then
                entidades(SafeCStr(item("IDComercial"))) = BuildEntidadForTablaPayload( _
                    SafeCStr(item("IDComercial")), SafeCStr(item("Comercial")), "Comercial")
            End If
        Next key
    End If

    If p_DTO.Exists("ColCPVs") Then
        Set col = p_DTO("ColCPVs")
        For Each key In col.Keys
            Set item = col(key)
            If Not item Is Nothing Then
                entidades(SafeCStr(item("IDCPV"))) = BuildEntidadForTablaPayload( _
                    SafeCStr(item("IDCPV")), SafeCStr(item("CPV")), "CPV")
            End If
        Next key
    End If

    If p_DTO.Exists("ColLugaresEjecucion") Then
        Set col = p_DTO("ColLugaresEjecucion")
        For Each key In col.Keys
            Set item = col(key)
            If Not item Is Nothing Then
                entidades(SafeCStr(item("IDLugarEjecucion"))) = BuildEntidadForTablaPayload( _
                    SafeCStr(item("IDLugarEjecucion")), SafeCStr(item("LugarEjecucion")), "Lugar")
            End If
        Next key
    End If

    If p_DTO.Exists("ColPECALES") Then
        Set col = p_DTO("ColPECALES")
        For Each key In col.Keys
            Set item = col(key)
            If Not item Is Nothing Then
                entidades(SafeCStr(item("IDPECAL"))) = BuildEntidadForTablaPayload( _
                    SafeCStr(item("IDPECAL")), SafeCStr(item("PECAL")), "PECAL")
            End If
        Next key
    End If

    If p_DTO.Exists("ColRACs") Then
        Set col = p_DTO("ColRACs")
        For Each key In col.Keys
            Set item = col(key)
            If Not item Is Nothing Then
                entidades(SafeCStr(item("IDRAC"))) = BuildEntidadForTablaPayload( _
                    SafeCStr(item("IDRAC")), SafeCStr(item("RAC")), "RAC")
            End If
        Next key
    End If

    If p_DTO.Exists("ColResponsables") Then
        Set col = p_DTO("ColResponsables")
        For Each key In col.Keys
            Set item = col(key)
            If Not item Is Nothing Then
                entidades(SafeCStr(item("IdUsuario"))) = BuildEntidadForTablaPayload( _
                    SafeCStr(item("IdUsuario")), SafeCStr(item("Nombre")), "Responsable")
            End If
        Next key
    End If

    Set BuildEntidadesPayloadFromDTO = entidades
End Function

Private Function BuildTablaDTO() As Object
    Dim dto As Object
    Set dto = CreateObject("Scripting.Dictionary")
    Set dto("Expediente") = BuildDTOExpediente(9999)
    Set dto("ColComerciales") = BuildDTOCollection("1", BuildTypedEntity("IDComercial", "1", "Comercial", "Comercial Uno"))
    Set dto("ColCPVs") = BuildDTOCollection("2", BuildTypedEntity("IDCPV", "2", "CPV", "CPV Dos"))
    Set dto("ColLugaresEjecucion") = BuildDTOCollection("3", BuildTypedEntity("IDLugarEjecucion", "3", "LugarEjecucion", "Lugar Tres"))
    Set dto("ColPECALES") = BuildDTOCollection("4", BuildTypedEntity("IDPECAL", "4", "PECAL", "PECAL Cuatro"))
    Set dto("ColRACs") = BuildDTOCollection("5", BuildTypedEntity("IDRAC", "5", "RAC", "RAC Cinco"))
    Set dto("ColResponsables") = BuildDTOCollection("6", BuildTypedEntity("IdUsuario", "6", "Nombre", "Responsable Seis"))
    Set BuildTablaDTO = dto
End Function

Private Sub AssertSerializedCount( _
    ByVal p_Tabla As Helper_ExpedienteEntidadesTabla, _
    ByVal p_ExpectedCount As Long, _
    ByVal p_ErrorContext As String _
)
    Dim p_Error As String
    Dim parsed As Object
    Set parsed = ParsedJson(p_Tabla.Serializar(p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CBool(parsed("ok")) <> True Then Err.Raise 1001, , p_ErrorContext & ": expected ok=true"
    If CLng(parsed("payload")("count")) <> p_ExpectedCount Then Err.Raise 1002, , p_ErrorContext & ": unexpected count"
End Sub

Public Function Test_Helper_ExpedienteEntidadesTabla_Init_StoresExpedienteID() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim tabla As New Helper_ExpedienteEntidadesTabla
    tabla.Init 9999, p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = ParsedJson(tabla.Serializar(p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CBool(parsed("ok")) <> True Then Err.Raise 1001, , "expected ok=true"
    If CLng(parsed("payload")("IDExpediente")) <> 9999 Then Err.Raise 1002, , "expected IDExpediente=9999"
    If CLng(parsed("payload")("count")) <> 0 Then Err.Raise 1003, , "expected empty initial state"

    Test_Helper_ExpedienteEntidadesTabla_Init_StoresExpedienteID = BuildOk("init-stores-id", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteEntidadesTabla_Init_StoresExpedienteID = BuildFail(Err.Description, logs)
End Function

Public Function Test_ExpedienteEntidades_CargarEntidadesTabla_WiresDtoStateIntoHelper() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim tabla As New Helper_ExpedienteEntidadesTabla
    Dim result As String
    Dim dto As Object
    Dim stage As String

    logs(0) = "setup"

    stage = "before-call"
    logs(1) = "building-dto"
    Set dto = BuildTablaDTO()

    logs(1) = "calling-helper"
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_CargarEntidadesTabla(tabla, dto, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    stage = "after-call"
    logs(2) = "parsing-result"
    Dim parsedResult As Object
    Set parsedResult = ParsedJson(result)
    If CBool(parsedResult("ok")) <> True Then Err.Raise 1001, , "expected helper envelope ok=true"
    If CLng(parsedResult("payload")("count")) <> 6 Then Err.Raise 1002, , "expected six entity-table items"

    stage = "before-serializar"
    Dim parsedState As Object
    Set parsedState = ParsedJson(tabla.Serializar(p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CLng(parsedState("payload")("IDExpediente")) <> 9999 Then Err.Raise 1003, , "expected state IDExpediente=9999"
    If CLng(parsedState("payload")("count")) <> 6 Then Err.Raise 1004, , "expected helper state count=6"
    If CStr(parsedState("payload")("entities")(1)("Tipo")) <> "Comercial" Then Err.Raise 1005, , "expected first entity type"

    Test_ExpedienteEntidades_CargarEntidadesTabla_WiresDtoStateIntoHelper = BuildOk("dto-state-wired", logs)
    Exit Function
EH:
    Test_ExpedienteEntidades_CargarEntidadesTabla_WiresDtoStateIntoHelper = BuildFail(stage & ": " & Err.Description, logs)
End Function

' Diagnostics helper: returns the exact envelope returned by the core loader.
' Useful when the main assertion-only test reports helper-level failure only.
Public Function Test_ExpedienteEntidades_CargarEntidadesTabla_WiresDtoStateIntoHelper_Diag() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim tabla As New Helper_ExpedienteEntidadesTabla
    Dim dto As Object
    Dim result As String

    logs(0) = "setup"
    Set dto = BuildTablaDTO()

    logs(1) = "calling-helper"
    result = modExpedienteEntidadesHelper.ExpedienteEntidades_CargarEntidadesTabla(tabla, dto)

    Test_ExpedienteEntidades_CargarEntidadesTabla_WiresDtoStateIntoHelper_Diag = result
    Exit Function

EH:
    Test_ExpedienteEntidades_CargarEntidadesTabla_WiresDtoStateIntoHelper_Diag = BuildFail(Err.Description, logs)
End Function

Public Function Test_Helper_ExpedienteEntidadesTabla_Serializar_UsesCanonicalJsonBoolean() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim tabla As New Helper_ExpedienteEntidadesTabla
    tabla.Init 9999, p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim rawJson As String
    rawJson = tabla.Serializar(p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If InStr(1, rawJson, """ok"":true", vbBinaryCompare) = 0 Then Err.Raise 1001, , "expected canonical JSON boolean true"

    Test_Helper_ExpedienteEntidadesTabla_Serializar_UsesCanonicalJsonBoolean = BuildOk("canonical-json-boolean", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteEntidadesTabla_Serializar_UsesCanonicalJsonBoolean = BuildFail(Err.Description, logs)
End Function

Private Sub AddEntityCollection( _
    ByVal p_Entidades As Object, _
    ByVal p_Col As Object, _
    ByVal p_IDProp As String, _
    ByVal p_NameProp As String, _
    ByVal p_Tipo As String _
)
    If p_Col Is Nothing Then Exit Sub
    Dim key As Variant
    Dim item As Object
    For Each key In p_Col.Keys
        Set item = p_Col(key)
        If Not item Is Nothing Then
            Set p_Entidades(CStr(item(p_IDProp))) = BuildEntity( _
                CLng(item(p_IDProp)), CStr(item(p_NameProp)), p_Tipo)
        End If
    Next key
End Sub

Public Function Test_Helper_ExpedienteEntidadesTabla_LoadsPayloadBuiltByDTO() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(8)
    On Error GoTo EH

    Dim dto As Object
    Dim entidades As Object
    Dim tabla As New Helper_ExpedienteEntidadesTabla
    Dim errMsg As String
    Dim payload As Object

    logs(0) = "setup"
    Set dto = BuildTablaDTO()

    logs(1) = "building-dict"
    Set entidades = CreateObject("Scripting.Dictionary")
    entidades.CompareMode = TextCompare

    AddEntityCollection entidades, dto("ColComerciales"), "IDComercial", "Comercial", "Comercial"
    AddEntityCollection entidades, dto("ColCPVs"), "IDCPV", "CPV", "CPV"
    AddEntityCollection entidades, dto("ColLugaresEjecucion"), "IDLugarEjecucion", "LugarEjecucion", "Lugar"
    AddEntityCollection entidades, dto("ColPECALES"), "IDPECAL", "PECAL", "PECAL"
    AddEntityCollection entidades, dto("ColRACs"), "IDRAC", "RAC", "RAC"
    AddEntityCollection entidades, dto("ColResponsables"), "IdUsuario", "Nombre", "Responsable"
    logs(2) = "count=" & entidades.Count

    logs(3) = "init"
    tabla.Init 9999, errMsg
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    logs(4) = "cargar"
    If tabla.Cargar(entidades, errMsg) <> True Then Err.Raise 1001, , errMsg

    logs(5) = "serializar"
    Set payload = ParsedJson(tabla.Serializar(errMsg))
    If CLng(payload("payload")("count")) <> 6 Then Err.Raise 1002, , "expected count=6"

    Test_Helper_ExpedienteEntidadesTabla_LoadsPayloadBuiltByDTO = BuildOk("full-test", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteEntidadesTabla_LoadsPayloadBuiltByDTO = BuildFail(Err.Description, logs)
End Function

Public Function Test_Helper_ExpedienteEntidadesTabla_Cargar_LoadsEntityData() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim tabla As New Helper_ExpedienteEntidadesTabla
    tabla.Init 9999, p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    If tabla.Cargar(BuildEntityCollection(), p_Error) <> True Then Err.Raise 1001, , "expected Cargar=True"
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = ParsedJson(tabla.Serializar(p_Error))
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If CLng(parsed("payload")("count")) <> 2 Then Err.Raise 1002, , "expected two loaded entities"
    If CStr(parsed("payload")("entities")(1)("Nombre")) <> "Comercial A" Then Err.Raise 1003, , "expected first entity data"
    If CStr(parsed("payload")("entities")(2)("Tipo")) <> "CPV" Then Err.Raise 1004, , "expected second entity type"

    Test_Helper_ExpedienteEntidadesTabla_Cargar_LoadsEntityData = BuildOk("cargar-loads-entity-data", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteEntidadesTabla_Cargar_LoadsEntityData = BuildFail(Err.Description, logs)
End Function

Public Function Test_Helper_ExpedienteEntidadesTabla_AgregarEliminar_HappyPathUpdatesState() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    On Error GoTo EH

    Dim p_Error As String
    Dim tabla As New Helper_ExpedienteEntidadesTabla
    tabla.Init 9999, p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    If tabla.AgregarEntidad(BuildEntity(1, "Comercial Uno"), p_Error) <> True Then Err.Raise 1001, , "expected add=true"
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    AssertSerializedCount tabla, 1, "after add"

    Dim parsed As Object
    Set parsed = ParsedJson(tabla.Serializar(p_Error))
    If CStr(parsed("payload")("entities")(1)("IDEntidad")) <> "1" Then Err.Raise 1002, , "expected IDEntidad=1"
    If CStr(parsed("payload")("entities")(1)("Nombre")) <> "Comercial Uno" Then Err.Raise 1003, , "expected entity name"

    If tabla.EliminarEntidad("1", p_Error) <> True Then Err.Raise 1004, , "expected delete=true"
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    AssertSerializedCount tabla, 0, "after delete"

    Test_Helper_ExpedienteEntidadesTabla_AgregarEliminar_HappyPathUpdatesState = BuildOk("add-delete-updates-state", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteEntidadesTabla_AgregarEliminar_HappyPathUpdatesState = BuildFail(Err.Description, logs)
End Function

Public Function Test_Helper_ExpedienteEntidadesTabla_AgregarEntidad_NothingReturnsError() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim tabla As New Helper_ExpedienteEntidadesTabla
    Dim entity As Object
    tabla.Init 9999, p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    On Error GoTo ExpectedError
    tabla.AgregarEntidad entity, p_Error
    Err.Raise 1001, , "expected Nothing entity to raise"

ExpectedError:
    On Error GoTo EH
    If Err.Number = 1001 Then Err.Raise 1001, , Err.Description
    If p_Error = "" Then Err.Raise 1002, , "expected p_Error for Nothing entity"
    If InStr(1, p_Error, "Nothing", vbTextCompare) = 0 Then Err.Raise 1003, , "expected Nothing error text"
    Test_Helper_ExpedienteEntidadesTabla_AgregarEntidad_NothingReturnsError = BuildOk("nothing-entity-error", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteEntidadesTabla_AgregarEntidad_NothingReturnsError = BuildFail(Err.Description, logs)
End Function

Public Function Test_Helper_ExpedienteEntidadesTabla_AgregarEntidad_MissingIDReturnsError() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim tabla As New Helper_ExpedienteEntidadesTabla
    tabla.Init 9999, p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    On Error GoTo ExpectedError
    tabla.AgregarEntidad BuildEntityWithoutID(), p_Error
    Err.Raise 1001, , "expected missing IDEntidad to raise"

ExpectedError:
    On Error GoTo EH
    If Err.Number = 1001 Then Err.Raise 1001, , Err.Description
    If p_Error = "" Then Err.Raise 1002, , "expected p_Error for missing IDEntidad"
    If InStr(1, p_Error, "IDEntidad", vbTextCompare) = 0 Then Err.Raise 1003, , "expected IDEntidad error text"
    Test_Helper_ExpedienteEntidadesTabla_AgregarEntidad_MissingIDReturnsError = BuildOk("missing-id-error", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteEntidadesTabla_AgregarEntidad_MissingIDReturnsError = BuildFail(Err.Description, logs)
End Function

Public Function Test_Helper_ExpedienteEntidadesTabla_DuplicateID_IsNoOp() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim tabla As New Helper_ExpedienteEntidadesTabla
    tabla.Init 9999, p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    If tabla.AgregarEntidad(BuildEntity(1, "Original"), p_Error) <> True Then Err.Raise 1001, , "expected first add=true"
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If tabla.AgregarEntidad(BuildEntity(1, "Replacement"), p_Error) <> False Then Err.Raise 1002, , "expected duplicate add=false"
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = ParsedJson(tabla.Serializar(p_Error))
    If CLng(parsed("payload")("count")) <> 1 Then Err.Raise 1003, , "expected duplicate to keep count=1"
    If CStr(parsed("payload")("entities")(1)("Nombre")) <> "Original" Then Err.Raise 1004, , "expected duplicate to preserve original entity"

    Test_Helper_ExpedienteEntidadesTabla_DuplicateID_IsNoOp = BuildOk("duplicate-is-no-op", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteEntidadesTabla_DuplicateID_IsNoOp = BuildFail(Err.Description, logs)
End Function

Public Function Test_Helper_ExpedienteEntidadesTabla_EliminarEntidad_MissingIDReturnsError() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim tabla As New Helper_ExpedienteEntidadesTabla
    tabla.Init 9999, p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    On Error GoTo ExpectedError
    tabla.EliminarEntidad "", p_Error
    Err.Raise 1001, , "expected missing IDEntidad to raise"

ExpectedError:
    On Error GoTo EH
    If Err.Number = 1001 Then Err.Raise 1001, , Err.Description
    If p_Error = "" Then Err.Raise 1002, , "expected p_Error for missing IDEntidad"
    If InStr(1, p_Error, "IDEntidad", vbTextCompare) = 0 Then Err.Raise 1003, , "expected IDEntidad error text"
    Test_Helper_ExpedienteEntidadesTabla_EliminarEntidad_MissingIDReturnsError = BuildOk("delete-missing-id-error", logs)
    Exit Function
EH:
    Test_Helper_ExpedienteEntidadesTabla_EliminarEntidad_MissingIDReturnsError = BuildFail(Err.Description, logs)
End Function

