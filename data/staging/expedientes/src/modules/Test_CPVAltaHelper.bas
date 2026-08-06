Attribute VB_Name = "Test_CPVAltaHelper"
Option Compare Database
Option Explicit

' Test_CPVAltaHelper — REWORK TDD atoms for modCPVAltaHelper.bas
' (Form_FormCPV Alta/Edición form, see docs/audit/cpv-alta-pure-data.md).

' === Module-level constants (all at top per vba-access §10.1) =======================

Private Const TEST_BASE_ID As Long = 900560

Private Const TEST_FIELD_NAME As String = "CPV"
Private Const TEST_FIELD_DESCRIPCION As String = "DESCRIPCION"


' === Local helpers (all at top per vba-access §10.1) ================================

Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

Private Function BuildStubValores( _
    ByVal p_Nombre As String, _
    ByVal p_Desc As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d(TEST_FIELD_NAME) = p_Nombre
    d(TEST_FIELD_DESCRIPCION) = p_Desc
    Set BuildStubValores = d
End Function

Private Function InsertRealRow( _
    ByVal p_ID As String, _
    ByVal p_Nombre As String, _
    ByVal p_Desc As String, _
    ByRef p_Error As String _
) As Boolean
    On Error GoTo EH
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim sql As String
    sql = "INSERT INTO TbCPV (IDCPV, CPV, Descripcion) VALUES (" & _
          p_ID & ", '" & Replace(p_Nombre, "'", "''") & "', '" & Replace(p_Desc, "'", "''") & "')"
    db.Execute sql, dbFailOnError
    InsertRealRow = True
    Exit Function
EH:
    p_Error = "InsertRealRow: " & Err.Description
    InsertRealRow = False
End Function

Private Function TeardownFixture() As Long
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim before As Long
    before = DCount("*", "TbCPV", "IDCPV >= " & TEST_BASE_ID)
    db.Execute "DELETE FROM TbCPV WHERE IDCPV >= " & TEST_BASE_ID, dbFailOnError
    On Error GoTo 0
    TeardownFixture = before
End Function


' === Public atoms ====================================================================

Public Function Test_CPVAltaHelper_Abrir_Inicializar_HappyAlta() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Set m_ObjCPVActivo = Nothing

    Dim p_Error As String
    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim json As String
    json = modCPVAltaHelper.CPVAlta_Abrir_Inicializar("alta", "", p_Entidad, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("titulo")) <> "ALTA DE CPV" Then Err.Raise 1001, , "expected titulo=ALTA DE CPV"
    If parsed("payload")("hasEntidad") <> False Then Err.Raise 1002, , "expected hasEntidad=false"

    Test_CPVAltaHelper_Abrir_Inicializar_HappyAlta = BuildOk("alta-initialized", logs)
    Exit Function
EH:
    Test_CPVAltaHelper_Abrir_Inicializar_HappyAlta = BuildFail(p_Error, logs)
End Function

Public Function Test_CPVAltaHelper_Abrir_Inicializar_HappyEdicion() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim seedId As String
    seedId = CStr(TEST_BASE_ID + 1)
    Dim insertErr As String
    If Not InsertRealRow(seedId, "EDIT-EXISTING", "edit-desc", insertErr) Then
        Err.Raise 1001, , insertErr
    End If

    Dim p_Error As String
    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim json As String
    json = modCPVAltaHelper.CPVAlta_Abrir_Inicializar("edicion", seedId, p_Entidad, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("titulo")) <> "EDICIÓN DE CPV" Then Err.Raise 1001, , "expected titulo=EDICIÓN DE CPV"
    If parsed("payload")("hasEntidad") <> True Then Err.Raise 1002, , "expected hasEntidad=true"
    If p_Entidad Is Nothing Then Err.Raise 1003, , "expected p_Entidad populated"
    If CStr(p_Entidad(TEST_FIELD_NAME)) <> "EDIT-EXISTING" Then Err.Raise 1004, , "expected CPV field populated"

    Test_CPVAltaHelper_Abrir_Inicializar_HappyEdicion = BuildOk("edicion-initialized", logs)
    Call TeardownFixture
    Set m_ObjCPVActivo = Nothing
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    Set m_ObjCPVActivo = Nothing
    On Error GoTo 0
    Test_CPVAltaHelper_Abrir_Inicializar_HappyEdicion = BuildFail(p_Error, logs)
End Function

Public Function Test_CPVAltaHelper_Abrir_Inicializar_SadInvalidModo() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Set m_ObjCPVActivo = Nothing

    Dim p_Error As String
    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim json As String
    json = modCPVAltaHelper.CPVAlta_Abrir_Inicializar("weird", "", p_Entidad, p_Error)
    If p_Error = "" Then Err.Raise 1001, , "expected p_Error populated for invalid modo"

    Test_CPVAltaHelper_Abrir_Inicializar_SadInvalidModo = BuildOk("invalid-modo-rejected", logs)
    Exit Function
EH:
    Test_CPVAltaHelper_Abrir_Inicializar_SadInvalidModo = BuildFail(p_Error, logs)
End Function

Public Function Test_CPVAltaHelper_Abrir_Inicializar_SadEdicionSinID() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Set m_ObjCPVActivo = Nothing

    Dim p_Error As String
    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim json As String
    json = modCPVAltaHelper.CPVAlta_Abrir_Inicializar("edicion", "", p_Entidad, p_Error)
    If p_Error = "" Then Err.Raise 1001, , "expected p_Error populated when edicion mode without ID"

    Test_CPVAltaHelper_Abrir_Inicializar_SadEdicionSinID = BuildOk("edicion-sin-id-rejected", logs)
    Exit Function
EH:
    Test_CPVAltaHelper_Abrir_Inicializar_SadEdicionSinID = BuildFail(p_Error, logs)
End Function

Public Function Test_CPVAltaHelper_VerificarCambios_HappyAltaAllEmpty() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("", "")

    Dim p_Error As String
    Dim json As String
    json = modCPVAltaHelper.CPVAlta_VerificarCambios(actuales, Nothing, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("hayCambios") <> False Then Err.Raise 1001, , "expected hayCambios=false for empty alta"

    Test_CPVAltaHelper_VerificarCambios_HappyAltaAllEmpty = BuildOk("alta-empty", logs)
    Exit Function
EH:
    Test_CPVAltaHelper_VerificarCambios_HappyAltaAllEmpty = BuildFail(p_Error, logs)
End Function

Public Function Test_CPVAltaHelper_VerificarCambios_HappyAltaOneField() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("SOMETHING", "")

    Dim p_Error As String
    Dim json As String
    json = modCPVAltaHelper.CPVAlta_VerificarCambios(actuales, Nothing, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("hayCambios") <> True Then Err.Raise 1001, , "expected hayCambios=true"

    Test_CPVAltaHelper_VerificarCambios_HappyAltaOneField = BuildOk("alta-one-field", logs)
    Exit Function
EH:
    Test_CPVAltaHelper_VerificarCambios_HappyAltaOneField = BuildFail(p_Error, logs)
End Function

Public Function Test_CPVAltaHelper_VerificarCambios_HappyEdicionAllEqual() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("SAME", "same-desc")
    Dim originales As Object
    Set originales = BuildStubValores("SAME", "same-desc")

    Dim p_Error As String
    Dim json As String
    json = modCPVAltaHelper.CPVAlta_VerificarCambios(actuales, originales, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("hayCambios") <> False Then Err.Raise 1001, , "expected hayCambios=false for equal fields"

    Test_CPVAltaHelper_VerificarCambios_HappyEdicionAllEqual = BuildOk("edicion-equal", logs)
    Exit Function
EH:
    Test_CPVAltaHelper_VerificarCambios_HappyEdicionAllEqual = BuildFail(p_Error, logs)
End Function

Public Function Test_CPVAltaHelper_VerificarCambios_HappyEdicionOneDiff() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("ORIGINAL", "edited-desc")
    Dim originales As Object
    Set originales = BuildStubValores("ORIGINAL", "original-desc")

    Dim p_Error As String
    Dim json As String
    json = modCPVAltaHelper.CPVAlta_VerificarCambios(actuales, originales, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("hayCambios") <> True Then Err.Raise 1001, , "expected hayCambios=true"
    If parsed("payload")("diffs")(TEST_FIELD_DESCRIPCION) <> True Then Err.Raise 1002, , "expected diffs.DESCRIPCION=true"

    Test_CPVAltaHelper_VerificarCambios_HappyEdicionOneDiff = BuildOk("edicion-one-diff", logs)
    Exit Function
EH:
    Test_CPVAltaHelper_VerificarCambios_HappyEdicionOneDiff = BuildFail(p_Error, logs)
End Function

Public Function Test_CPVAltaHelper_VerificarCambios_SadNoActuales() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modCPVAltaHelper.CPVAlta_VerificarCambios(Nothing, Nothing, p_Error)
    If p_Error = "" Then Err.Raise 1001, , "expected p_Error populated for Nothing actuales"

    Test_CPVAltaHelper_VerificarCambios_SadNoActuales = BuildOk("no-actuales-rejected", logs)
    Exit Function
EH:
    Test_CPVAltaHelper_VerificarCambios_SadNoActuales = BuildFail(p_Error, logs)
End Function

Public Function Test_CPVAltaHelper_Registrar_HappyAlta() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    Set m_ObjCPVActivo = Nothing

    Dim valores As Object
    Set valores = BuildStubValores("ALTA-TEST-CPV", "alta-desc")

    Dim p_Error As String
    Dim json As String
    json = modCPVAltaHelper.CPVAlta_Registrar(valores, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("ok") <> True Then Err.Raise 1001, , "expected ok=true"
    If CStr(parsed("payload")("modo")) <> "alta" Then Err.Raise 1002, , "expected modo=alta"

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim countAfter As Long
    countAfter = DCount("*", "TbCPV", "CPV='ALTA-TEST-CPV'")
    If countAfter <> 1 Then Err.Raise 1003, , "expected 1 row after alta"

    Test_CPVAltaHelper_Registrar_HappyAlta = BuildOk("alta", logs)
    Call TeardownFixture
    Set m_ObjCPVActivo = Nothing
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    Set m_ObjCPVActivo = Nothing
    On Error GoTo 0
    Test_CPVAltaHelper_Registrar_HappyAlta = BuildFail(p_Error, logs)
End Function

Public Function Test_CPVAltaHelper_Registrar_HappyEdicion() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    Dim seedId As String
    seedId = CStr(TEST_BASE_ID + 1)
    Dim insertErr As String
    If Not InsertRealRow(seedId, "ORIGINAL-CPV", "orig-desc", insertErr) Then
        Err.Raise 1001, , insertErr
    End If

    Dim active As New CPV
    active.IDCPV = seedId
    Set m_ObjCPVActivo = active

    Dim valores As Object
    Set valores = BuildStubValores("EDITED-CPV", "edited-desc")

    Dim p_Error As String
    Dim json As String
    json = modCPVAltaHelper.CPVAlta_Registrar(valores, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("ok") <> True Then Err.Raise 1001, , "expected ok=true"
    If CStr(parsed("payload")("modo")) <> "edicion" Then Err.Raise 1002, , "expected modo=edicion"

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim countAfter As Long
    countAfter = DCount("*", "TbCPV", "IDCPV=" & seedId & " AND CPV='EDITED-CPV'")
    If countAfter <> 1 Then Err.Raise 1003, , "expected 1 row with new CPV value"

    Test_CPVAltaHelper_Registrar_HappyEdicion = BuildOk("edicion", logs)
    Call TeardownFixture
    Set m_ObjCPVActivo = Nothing
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    Set m_ObjCPVActivo = Nothing
    On Error GoTo 0
    Test_CPVAltaHelper_Registrar_HappyEdicion = BuildFail(p_Error, logs)
End Function

Public Function Test_CPVAltaHelper_Registrar_SadNoValores() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Set m_ObjCPVActivo = Nothing

    Dim p_Error As String
    Dim json As String
    json = modCPVAltaHelper.CPVAlta_Registrar(Nothing, p_Error)
    If p_Error = "" Then Err.Raise 1001, , "expected p_Error populated for Nothing valores"

    Test_CPVAltaHelper_Registrar_SadNoValores = BuildOk("no-valores-rejected", logs)
    Set m_ObjCPVActivo = Nothing
    Exit Function
EH:
    On Error Resume Next
    Set m_ObjCPVActivo = Nothing
    On Error GoTo 0
    Test_CPVAltaHelper_Registrar_SadNoValores = BuildFail(p_Error, logs)
End Function

Public Function Test_CPVAltaHelper_Cerrar_HappyCleared() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim preObj As New CPV
    preObj.IDCPV = "999999"
    Set m_ObjCPVActivo = preObj

    Dim p_Error As String
    Dim json As String
    json = modCPVAltaHelper.CPVAlta_Cerrar(p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    If Not m_ObjCPVActivo Is Nothing Then Err.Raise 1001, , "expected m_ObjCPVActivo cleared"

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("ok") <> True Then Err.Raise 1002, , "expected payload.ok=true"

    Test_CPVAltaHelper_Cerrar_HappyCleared = BuildOk("cleared", logs)
    Exit Function
EH:
    Test_CPVAltaHelper_Cerrar_HappyCleared = BuildFail(p_Error, logs)
End Function

Public Function Test_CPVAltaHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_CPVAltaHelper_Abrir_Inicializar_HappyAlta", _
        "Test_CPVAltaHelper_Abrir_Inicializar_HappyEdicion", _
        "Test_CPVAltaHelper_Abrir_Inicializar_SadInvalidModo", _
        "Test_CPVAltaHelper_Abrir_Inicializar_SadEdicionSinID", _
        "Test_CPVAltaHelper_VerificarCambios_HappyAltaAllEmpty", _
        "Test_CPVAltaHelper_VerificarCambios_HappyAltaOneField", _
        "Test_CPVAltaHelper_VerificarCambios_HappyEdicionAllEqual", _
        "Test_CPVAltaHelper_VerificarCambios_HappyEdicionOneDiff", _
        "Test_CPVAltaHelper_VerificarCambios_SadNoActuales", _
        "Test_CPVAltaHelper_Registrar_HappyAlta", _
        "Test_CPVAltaHelper_Registrar_HappyEdicion", _
        "Test_CPVAltaHelper_Registrar_SadNoValores", _
        "Test_CPVAltaHelper_Cerrar_HappyCleared" _
    )

    Dim passed As Long
    Dim failed As Long
    Dim firstFailure As String
    Dim i As Long
    For i = LBound(atoms) To UBound(atoms)
        Dim result As String
        result = Application.Run(CStr(atoms(i)))
        If InStr(result, """ok"":true") > 0 Then
            passed = passed + 1
        Else
            failed = failed + 1
            If Len(firstFailure) = 0 Then firstFailure = CStr(atoms(i)) & " -> " & result
        End If
    Next i

    logs(0) = "RunAll: passed=" & passed & " failed=" & failed
    If failed > 0 Then
        logs(1) = "RunAll: firstFailure=" & firstFailure
        Test_CPVAltaHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_CPVAltaHelper_RunAll = BuildOk("all-passed", logs)
    End If
    On Error Resume Next
    Call TeardownFixture
    Set m_ObjCPVActivo = Nothing
    On Error GoTo 0
    Exit Function
EH:
    On Error Resume Next
    Dim p_Error As String
    p_Error = "Test_CPVAltaHelper_RunAll EH: " & Err.Description
    Call TeardownFixture
    Set m_ObjCPVActivo = Nothing
    On Error GoTo 0
    Test_CPVAltaHelper_RunAll = BuildFail(p_Error, logs)
End Function
