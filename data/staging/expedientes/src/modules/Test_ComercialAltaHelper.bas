Attribute VB_Name = "Test_ComercialAltaHelper"
Option Compare Database
Option Explicit

' Test_ComercialAltaHelper — REWORK TDD atoms for modComercialAltaHelper.bas
' (Form_FormComercial Alta/Edición form, see docs/audit/comercial-alta-pure-data.md).

' === Module-level constants (all at top per vba-access §10.1) =======================

Private Const TEST_BASE_ID As Long = 900550

Private Const TEST_FIELD_NAME As String = "Comercial"
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
    sql = "INSERT INTO TbComerciales (IDComercial, Comercial, Descripcion) VALUES (" & _
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
    before = DCount("*", "TbComerciales", "IDComercial >= " & TEST_BASE_ID)
    db.Execute "DELETE FROM TbComerciales WHERE IDComercial >= " & TEST_BASE_ID, dbFailOnError
    On Error GoTo 0
    TeardownFixture = before
End Function


' === Public atoms ====================================================================

' 1. Abrir_Inicializar — happy alta
Public Function Test_ComercialAltaHelper_Abrir_Inicializar_HappyAlta() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Set m_ObjComercialActivo = Nothing

    Dim p_Error As String
    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim json As String
    json = modComercialAltaHelper.ComercialAlta_Abrir_Inicializar("alta", "", p_Entidad, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("titulo")) <> "ALTA DE COMERCIAL" Then Err.Raise 1001, , "expected titulo=ALTA DE COMERCIAL"
    If CStr(parsed("payload")("modo")) <> "alta" Then Err.Raise 1002, , "expected modo=alta"
    If parsed("payload")("hasEntidad") <> False Then Err.Raise 1003, , "expected hasEntidad=false"
    If p_Entidad Is Nothing Then Err.Raise 1004, , "expected p_Entidad populated as empty stub"

    Test_ComercialAltaHelper_Abrir_Inicializar_HappyAlta = BuildOk("alta-initialized", logs)
    Exit Function
EH:
    Test_ComercialAltaHelper_Abrir_Inicializar_HappyAlta = BuildFail(p_Error, logs)
End Function

' 2. Abrir_Inicializar — happy edicion
Public Function Test_ComercialAltaHelper_Abrir_Inicializar_HappyEdicion() As String
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
    json = modComercialAltaHelper.ComercialAlta_Abrir_Inicializar("edicion", seedId, p_Entidad, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("titulo")) <> "EDICIÓN DE COMERCIAL" Then Err.Raise 1001, , "expected titulo=EDICIÓN DE COMERCIAL"
    If parsed("payload")("hasEntidad") <> True Then Err.Raise 1002, , "expected hasEntidad=true"
    If p_Entidad Is Nothing Then Err.Raise 1003, , "expected p_Entidad populated"
    If CStr(p_Entidad(TEST_FIELD_NAME)) <> "EDIT-EXISTING" Then Err.Raise 1004, , "expected Comercial field populated"

    Test_ComercialAltaHelper_Abrir_Inicializar_HappyEdicion = BuildOk("edicion-initialized", logs)
    Call TeardownFixture
    Set m_ObjComercialActivo = Nothing
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    Set m_ObjComercialActivo = Nothing
    On Error GoTo 0
    Test_ComercialAltaHelper_Abrir_Inicializar_HappyEdicion = BuildFail(p_Error, logs)
End Function

' 3. Abrir_Inicializar — sad invalid modo
Public Function Test_ComercialAltaHelper_Abrir_Inicializar_SadInvalidModo() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Set m_ObjComercialActivo = Nothing

    Dim p_Error As String
    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim json As String
    json = modComercialAltaHelper.ComercialAlta_Abrir_Inicializar("weird", "", p_Entidad, p_Error)
    If p_Error = "" Then Err.Raise 1001, , "expected p_Error populated for invalid modo"
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false"

    Test_ComercialAltaHelper_Abrir_Inicializar_SadInvalidModo = BuildOk("invalid-modo-rejected", logs)
    Exit Function
EH:
    Test_ComercialAltaHelper_Abrir_Inicializar_SadInvalidModo = BuildFail(p_Error, logs)
End Function

' 4. Abrir_Inicializar — sad edicion sin ID
Public Function Test_ComercialAltaHelper_Abrir_Inicializar_SadEdicionSinID() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Set m_ObjComercialActivo = Nothing

    Dim p_Error As String
    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim json As String
    json = modComercialAltaHelper.ComercialAlta_Abrir_Inicializar("edicion", "", p_Entidad, p_Error)
    If p_Error = "" Then Err.Raise 1001, , "expected p_Error populated when edicion mode without ID"
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false"

    Test_ComercialAltaHelper_Abrir_Inicializar_SadEdicionSinID = BuildOk("edicion-sin-id-rejected", logs)
    Exit Function
EH:
    Test_ComercialAltaHelper_Abrir_Inicializar_SadEdicionSinID = BuildFail(p_Error, logs)
End Function

' 5. VerificarCambios — alta with all-empty values -> hayCambios=false
Public Function Test_ComercialAltaHelper_VerificarCambios_HappyAltaAllEmpty() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("", "")

    Dim p_Error As String
    Dim json As String
    json = modComercialAltaHelper.ComercialAlta_VerificarCambios(actuales, Nothing, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("hayCambios") <> False Then Err.Raise 1001, , "expected hayCambios=false for empty alta"

    Test_ComercialAltaHelper_VerificarCambios_HappyAltaAllEmpty = BuildOk("alta-empty", logs)
    Exit Function
EH:
    Test_ComercialAltaHelper_VerificarCambios_HappyAltaAllEmpty = BuildFail(p_Error, logs)
End Function

' 6. VerificarCambios — alta with one field filled -> hayCambios=true
Public Function Test_ComercialAltaHelper_VerificarCambios_HappyAltaOneField() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("SOMETHING", "")

    Dim p_Error As String
    Dim json As String
    json = modComercialAltaHelper.ComercialAlta_VerificarCambios(actuales, Nothing, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("hayCambios") <> True Then Err.Raise 1001, , "expected hayCambios=true"

    Test_ComercialAltaHelper_VerificarCambios_HappyAltaOneField = BuildOk("alta-one-field", logs)
    Exit Function
EH:
    Test_ComercialAltaHelper_VerificarCambios_HappyAltaOneField = BuildFail(p_Error, logs)
End Function

' 7. VerificarCambios — edicion all equal -> hayCambios=false
Public Function Test_ComercialAltaHelper_VerificarCambios_HappyEdicionAllEqual() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("SAME", "same-desc")
    Dim originales As Object
    Set originales = BuildStubValores("SAME", "same-desc")

    Dim p_Error As String
    Dim json As String
    json = modComercialAltaHelper.ComercialAlta_VerificarCambios(actuales, originales, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("hayCambios") <> False Then Err.Raise 1001, , "expected hayCambios=false for equal fields"

    Test_ComercialAltaHelper_VerificarCambios_HappyEdicionAllEqual = BuildOk("edicion-equal", logs)
    Exit Function
EH:
    Test_ComercialAltaHelper_VerificarCambios_HappyEdicionAllEqual = BuildFail(p_Error, logs)
End Function

' 8. VerificarCambios — edicion one field differs -> hayCambios=true
Public Function Test_ComercialAltaHelper_VerificarCambios_HappyEdicionOneDiff() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("ORIGINAL", "edited-desc")
    Dim originales As Object
    Set originales = BuildStubValores("ORIGINAL", "original-desc")

    Dim p_Error As String
    Dim json As String
    json = modComercialAltaHelper.ComercialAlta_VerificarCambios(actuales, originales, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("hayCambios") <> True Then Err.Raise 1001, , "expected hayCambios=true"
    If parsed("payload")("diffs")(TEST_FIELD_DESCRIPCION) <> True Then Err.Raise 1002, , "expected diffs.DESCRIPCION=true"

    Test_ComercialAltaHelper_VerificarCambios_HappyEdicionOneDiff = BuildOk("edicion-one-diff", logs)
    Exit Function
EH:
    Test_ComercialAltaHelper_VerificarCambios_HappyEdicionOneDiff = BuildFail(p_Error, logs)
End Function

' 9. VerificarCambios — sad no actuales
Public Function Test_ComercialAltaHelper_VerificarCambios_SadNoActuales() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modComercialAltaHelper.ComercialAlta_VerificarCambios(Nothing, Nothing, p_Error)
    If p_Error = "" Then Err.Raise 1001, , "expected p_Error populated for Nothing actuales"
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false"

    Test_ComercialAltaHelper_VerificarCambios_SadNoActuales = BuildOk("no-actuales-rejected", logs)
    Exit Function
EH:
    Test_ComercialAltaHelper_VerificarCambios_SadNoActuales = BuildFail(p_Error, logs)
End Function

' 10. Registrar — happy alta (m_ObjComercialActivo is Nothing -> insert)
Public Function Test_ComercialAltaHelper_Registrar_HappyAlta() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    Set m_ObjComercialActivo = Nothing

    Dim valores As Object
    Set valores = BuildStubValores("ALTA-TEST-COMERCIAL", "alta-desc")

    Dim p_Error As String
    Dim json As String
    json = modComercialAltaHelper.ComercialAlta_Registrar(valores, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("ok") <> True Then Err.Raise 1001, , "expected ok=true"
    If CStr(parsed("payload")("modo")) <> "alta" Then Err.Raise 1002, , "expected modo=alta"

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim countAfter As Long
    countAfter = DCount("*", "TbComerciales", "Comercial='ALTA-TEST-COMERCIAL'")
    If countAfter <> 1 Then Err.Raise 1003, , "expected 1 row after alta, got " & countAfter

    Test_ComercialAltaHelper_Registrar_HappyAlta = BuildOk("alta", logs)
    Call TeardownFixture
    Set m_ObjComercialActivo = Nothing
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    Set m_ObjComercialActivo = Nothing
    On Error GoTo 0
    Test_ComercialAltaHelper_Registrar_HappyAlta = BuildFail(p_Error, logs)
End Function

' 11. Registrar — happy edicion (m_ObjComercialActivo set -> update)
Public Function Test_ComercialAltaHelper_Registrar_HappyEdicion() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    Dim seedId As String
    seedId = CStr(TEST_BASE_ID + 1)
    Dim insertErr As String
    If Not InsertRealRow(seedId, "ORIGINAL-COMERCIAL", "orig-desc", insertErr) Then
        Err.Raise 1001, , insertErr
    End If

    Dim active As New Comercial
    active.IDComercial = seedId
    Set m_ObjComercialActivo = active

    Dim valores As Object
    Set valores = BuildStubValores("EDITED-COMERCIAL", "edited-desc")

    Dim p_Error As String
    Dim json As String
    json = modComercialAltaHelper.ComercialAlta_Registrar(valores, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("ok") <> True Then Err.Raise 1001, , "expected ok=true"
    If CStr(parsed("payload")("modo")) <> "edicion" Then Err.Raise 1002, , "expected modo=edicion"

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim countAfter As Long
    countAfter = DCount("*", "TbComerciales", "IDComercial=" & seedId & " AND Comercial='EDITED-COMERCIAL'")
    If countAfter <> 1 Then Err.Raise 1003, , "expected 1 row with new Comercial value"

    Test_ComercialAltaHelper_Registrar_HappyEdicion = BuildOk("edicion", logs)
    Call TeardownFixture
    Set m_ObjComercialActivo = Nothing
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    Set m_ObjComercialActivo = Nothing
    On Error GoTo 0
    Test_ComercialAltaHelper_Registrar_HappyEdicion = BuildFail(p_Error, logs)
End Function

' 12. Registrar — sad no valores
Public Function Test_ComercialAltaHelper_Registrar_SadNoValores() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Set m_ObjComercialActivo = Nothing

    Dim p_Error As String
    Dim json As String
    json = modComercialAltaHelper.ComercialAlta_Registrar(Nothing, p_Error)
    If p_Error = "" Then Err.Raise 1001, , "expected p_Error populated for Nothing valores"
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false"

    Test_ComercialAltaHelper_Registrar_SadNoValores = BuildOk("no-valores-rejected", logs)
    Set m_ObjComercialActivo = Nothing
    Exit Function
EH:
    On Error Resume Next
    Set m_ObjComercialActivo = Nothing
    On Error GoTo 0
    Test_ComercialAltaHelper_Registrar_SadNoValores = BuildFail(p_Error, logs)
End Function

' 13. Cerrar — happy (m_ObjComercialActivo cleared, ok=true)
Public Function Test_ComercialAltaHelper_Cerrar_HappyCleared() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim preObj As New Comercial
    preObj.IDComercial = "999999"
    Set m_ObjComercialActivo = preObj

    Dim p_Error As String
    Dim json As String
    json = modComercialAltaHelper.ComercialAlta_Cerrar(p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    If Not m_ObjComercialActivo Is Nothing Then Err.Raise 1001, , "expected m_ObjComercialActivo cleared"

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("ok") <> True Then Err.Raise 1002, , "expected payload.ok=true"

    Test_ComercialAltaHelper_Cerrar_HappyCleared = BuildOk("cleared", logs)
    Exit Function
EH:
    Test_ComercialAltaHelper_Cerrar_HappyCleared = BuildFail(p_Error, logs)
End Function

' 14. RunAll — wrapper for Dysflow manifest discovery
Public Function Test_ComercialAltaHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_ComercialAltaHelper_Abrir_Inicializar_HappyAlta", _
        "Test_ComercialAltaHelper_Abrir_Inicializar_HappyEdicion", _
        "Test_ComercialAltaHelper_Abrir_Inicializar_SadInvalidModo", _
        "Test_ComercialAltaHelper_Abrir_Inicializar_SadEdicionSinID", _
        "Test_ComercialAltaHelper_VerificarCambios_HappyAltaAllEmpty", _
        "Test_ComercialAltaHelper_VerificarCambios_HappyAltaOneField", _
        "Test_ComercialAltaHelper_VerificarCambios_HappyEdicionAllEqual", _
        "Test_ComercialAltaHelper_VerificarCambios_HappyEdicionOneDiff", _
        "Test_ComercialAltaHelper_VerificarCambios_SadNoActuales", _
        "Test_ComercialAltaHelper_Registrar_HappyAlta", _
        "Test_ComercialAltaHelper_Registrar_HappyEdicion", _
        "Test_ComercialAltaHelper_Registrar_SadNoValores", _
        "Test_ComercialAltaHelper_Cerrar_HappyCleared" _
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
        Test_ComercialAltaHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_ComercialAltaHelper_RunAll = BuildOk("all-passed", logs)
    End If
    On Error Resume Next
    Call TeardownFixture
    Set m_ObjComercialActivo = Nothing
    On Error GoTo 0
    Exit Function
EH:
    On Error Resume Next
    Dim p_Error As String
    p_Error = "Test_ComercialAltaHelper_RunAll EH: " & Err.Description
    Call TeardownFixture
    Set m_ObjComercialActivo = Nothing
    On Error GoTo 0
    Test_ComercialAltaHelper_RunAll = BuildFail(p_Error, logs)
End Function
