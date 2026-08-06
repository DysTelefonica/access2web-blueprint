Attribute VB_Name = "Test_RACHelper"
Option Compare Database
Option Explicit

' Test_RACHelper — REWORK TDD atoms for modRACHelper.bas
' (Form_FormRAC Alta/Edición form, see docs/audit/rac-thin.md).
'
' Anti-pattern removed (was in PR #29 commit 979a57a):
'   - Test_RACHelper_OpenForm(ByRef p_Form, ByRef p_Error) used to call
'     DoCmd.OpenForm TEST_FORM_NAME and return Forms(TEST_FORM_NAME) as p_Form.
'   - Every atom opened the form via OpenForm, populated controls via
'     p_Form.Controls("X").Value = ..., and asserted on p_Form.Controls(...) state.
'
' New design (per access-vba-e2e-methodology rule #1):
'   - ZERO DoCmd.OpenForm calls in tests.
'   - ZERO Forms(...) references in tests.
'   - ZERO Screen.ActiveForm references in tests.
'   - ZERO Application.Echo in tests.
'   - Stubs are Scripting.Dictionary instances, NOT real form-bound objects.
'   - Helpers receive pure data; atoms parse JSON via JsonConverter.
'
' Implements access-vba-tdd skill:
'   - §1.8 declaration ordering.
'   - §1.10 atom signature MUST match helper signature EXACTLY.
'   - §4.2 no-humo: atoms assert concrete values, not "did not crash".
'   - §4.4 strong assertions: count=0 vs >0, payload keys present, etc.
'   - §4.5 cardinalidad for mutaciones: countBefore / countAfter for Eliminar tests.

' === Module-level constants (all at top per vba-access §10.1) =======================

' Fixture ID base — entities created by atoms use IDs in this range so teardown is
' surgical (DELETE WHERE IDRAC >= TEST_BASE_ID).
Private Const TEST_BASE_ID As Long = 900900


' === Local helpers (all at top per vba-access §10.1) ================================

' --- BuildOk / BuildFail --------------------------------------------------------------
' Thin wrappers around TestingCore_BuildOk / BuildFail so atoms read like:
'   Test_X = BuildOk(...)
Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

' --- BuildStubValores ----------------------------------------------------------------
' Builds a Dictionary {RAC, CORREO, DESCRIPCION} representing form values.
' Used to pass into the helper without depending on a real form.
Private Function BuildStubValores( _
    ByVal p_RAC As String, _
    ByVal p_Correo As String, _
    ByVal p_Desc As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RAC") = p_RAC
    d("CORREO") = p_Correo
    d("DESCRIPCION") = p_Desc
    Set BuildStubValores = d
End Function

' --- InsertRealRow -------------------------------------------------------------------
' Inserts a row into TbRACS for the Registrar_HappyEdicion test.
Private Function InsertRealRow( _
    ByVal p_ID As String, _
    ByVal p_RAC As String, _
    ByVal p_Correo As String, _
    ByVal p_Desc As String, _
    ByRef p_Error As String _
) As Boolean
    On Error GoTo EH
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim sql As String
    sql = "INSERT INTO TbRACS (IDRAC, RAC, CORREO, DESCRIPCION) VALUES (" & _
          p_ID & ", '" & Replace(p_RAC, "'", "''") & "', '" & Replace(p_Correo, "'", "''") & "', '" & Replace(p_Desc, "'", "''") & "')"
    db.Execute sql, dbFailOnError
    InsertRealRow = True
    Exit Function
EH:
    p_Error = "InsertRealRow: " & Err.Description
    InsertRealRow = False
End Function

' --- TeardownFixture ------------------------------------------------------------------
' Deletes any rows whose ID is in the test range. Returns the count deleted.
Private Function TeardownFixture() As Long
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim before As Long
    before = DCount("*", "TbRACS", "IDRAC >= " & TEST_BASE_ID)
    db.Execute "DELETE FROM TbRACS WHERE IDRAC >= " & TEST_BASE_ID, dbFailOnError
    On Error GoTo 0
    TeardownFixture = before
End Function

' --- AssertPayloadHasKey --------------------------------------------------------------
' Asserts that parsed(payload) is a Dictionary and contains p_Key. Returns True if OK.
' Returns False (and writes to p_Message) if not. Split guard per vba-access §1.6.1.
Private Function AssertPayloadHasKey( _
    ByVal p_Parsed As Object, _
    ByVal p_Key As String, _
    ByRef p_Message As String _
) As Boolean
    If p_Parsed Is Nothing Then
        p_Message = "parsed payload is Nothing"
        AssertPayloadHasKey = False
        Exit Function
    End If
    If Not p_Parsed.Exists(p_Key) Then
        p_Message = "payload missing key '" & p_Key & "'"
        AssertPayloadHasKey = False
        Exit Function
    End If
    AssertPayloadHasKey = True
End Function


' === Public atoms ====================================================================

' ---------------------------------------------------------------------------
' 1. Abrir_Inicializar — happy alta (no ID, no m_ObjRACActivo)
' ---------------------------------------------------------------------------
Public Function Test_RACHelper_Abrir_Inicializar_HappyAlta() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    ' Ensure m_ObjRACActivo is clean.
    Set m_ObjRACActivo = Nothing

    Dim p_Error As String
    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim json As String
    json = modRACHelper.RAC_Abrir_Inicializar("alta", "", p_Entidad, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If InStr(json, """ok"":true") = 0 Then Err.Raise 1001, , "expected ok=true, got " & json

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim msg As String
    If Not AssertPayloadHasKey(parsed("payload"), "titulo", msg) Then Err.Raise 1002, , msg
    If Not AssertPayloadHasKey(parsed("payload"), "modo", msg) Then Err.Raise 1002, , msg
    If Not AssertPayloadHasKey(parsed("payload"), "hasEntidad", msg) Then Err.Raise 1002, , msg

    If CStr(parsed("payload")("titulo")) <> "ALTA DE RAC" Then Err.Raise 1003, , "expected titulo=ALTA DE RAC"
    If CStr(parsed("payload")("modo")) <> "alta" Then Err.Raise 1004, , "expected modo=alta"
    If parsed("payload")("hasEntidad") <> False Then Err.Raise 1005, , "expected hasEntidad=false"
    If p_Entidad Is Nothing Then Err.Raise 1006, , "expected p_Entidad populated as empty stub"

    Test_RACHelper_Abrir_Inicializar_HappyAlta = BuildOk("alta-initialized", logs)
    Exit Function
EH:
    Test_RACHelper_Abrir_Inicializar_HappyAlta = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 2. Abrir_Inicializar — happy edicion (seeded row + ID)
' ---------------------------------------------------------------------------
Public Function Test_RACHelper_Abrir_Inicializar_HappyEdicion() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim seedId As String
    seedId = CStr(TEST_BASE_ID + 1)
    Dim insertErr As String
    If Not InsertRealRow(seedId, "EDIT-EXISTING", "edit@example.com", "edit-desc", insertErr) Then
        Err.Raise 1001, , insertErr
    End If

    Dim p_Error As String
    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim json As String
    json = modRACHelper.RAC_Abrir_Inicializar("edicion", seedId, p_Entidad, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("titulo")) <> "EDICIÓN DE RAC" Then Err.Raise 1002, , "expected titulo=EDICIÓN DE RAC"
    If CStr(parsed("payload")("modo")) <> "edicion" Then Err.Raise 1003, , "expected modo=edicion"
    If parsed("payload")("hasEntidad") <> True Then Err.Raise 1004, , "expected hasEntidad=true"
    If CStr(parsed("payload")("id")) <> seedId Then Err.Raise 1005, , "expected id=" & seedId
    If p_Entidad Is Nothing Then Err.Raise 1006, , "expected p_Entidad populated"
    If CStr(p_Entidad("RAC")) <> "EDIT-EXISTING" Then Err.Raise 1007, , "expected RAC field populated"
    If CStr(p_Entidad("CORREO")) <> "edit@example.com" Then Err.Raise 1008, , "expected CORREO field populated"

    Test_RACHelper_Abrir_Inicializar_HappyEdicion = BuildOk("edicion-initialized", logs)
    Call TeardownFixture
    Set m_ObjRACActivo = Nothing
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    Set m_ObjRACActivo = Nothing
    On Error GoTo 0
    Test_RACHelper_Abrir_Inicializar_HappyEdicion = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 3. Abrir_Inicializar — sad invalid modo
' ---------------------------------------------------------------------------
Public Function Test_RACHelper_Abrir_Inicializar_SadInvalidModo() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Set m_ObjRACActivo = Nothing

    Dim p_Error As String
    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim json As String
    json = modRACHelper.RAC_Abrir_Inicializar("weird", "", p_Entidad, p_Error)
    If p_Error = "" Then Err.Raise 1001, , "expected p_Error populated for invalid modo"
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false, got " & json

    Test_RACHelper_Abrir_Inicializar_SadInvalidModo = BuildOk("invalid-modo-rejected", logs)
    Exit Function
EH:
    Test_RACHelper_Abrir_Inicializar_SadInvalidModo = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 4. Abrir_Inicializar — sad edicion without ID
' ---------------------------------------------------------------------------
Public Function Test_RACHelper_Abrir_Inicializar_SadEdicionSinID() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Set m_ObjRACActivo = Nothing

    Dim p_Error As String
    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim json As String
    json = modRACHelper.RAC_Abrir_Inicializar("edicion", "", p_Entidad, p_Error)
    If p_Error = "" Then Err.Raise 1001, , "expected p_Error populated when edicion mode without ID"
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false, got " & json

    Test_RACHelper_Abrir_Inicializar_SadEdicionSinID = BuildOk("edicion-sin-id-rejected", logs)
    Exit Function
EH:
    Test_RACHelper_Abrir_Inicializar_SadEdicionSinID = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 5. Verificar_Cambios — alta with all-empty values -> hayCambios=false
' ---------------------------------------------------------------------------
Public Function Test_RACHelper_Verificar_Cambios_HappyAltaAllEmpty() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("", "", "")

    Dim p_Error As String
    Dim json As String
    json = modRACHelper.RAC_Verificar_Cambios(actuales, Nothing, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hayCambios") <> False Then Err.Raise 1002, , "expected hayCambios=false for empty alta"

    Test_RACHelper_Verificar_Cambios_HappyAltaAllEmpty = BuildOk("alta-empty", logs)
    Exit Function
EH:
    Test_RACHelper_Verificar_Cambios_HappyAltaAllEmpty = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 6. Verificar_Cambios — alta with one field filled -> hayCambios=true
' ---------------------------------------------------------------------------
Public Function Test_RACHelper_Verificar_Cambios_HappyAltaOneField() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("SOMETHING", "", "")

    Dim p_Error As String
    Dim json As String
    json = modRACHelper.RAC_Verificar_Cambios(actuales, Nothing, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hayCambios") <> True Then Err.Raise 1002, , "expected hayCambios=true for alta with content"

    Test_RACHelper_Verificar_Cambios_HappyAltaOneField = BuildOk("alta-one-field", logs)
    Exit Function
EH:
    Test_RACHelper_Verificar_Cambios_HappyAltaOneField = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 7. Verificar_Cambios — edicion all equal -> hayCambios=false
' ---------------------------------------------------------------------------
Public Function Test_RACHelper_Verificar_Cambios_HappyEdicionAllEqual() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("SAME", "same@example.com", "same-desc")
    Dim originales As Object
    Set originales = BuildStubValores("SAME", "same@example.com", "same-desc")

    Dim p_Error As String
    Dim json As String
    json = modRACHelper.RAC_Verificar_Cambios(actuales, originales, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hayCambios") <> False Then Err.Raise 1002, , "expected hayCambios=false for equal fields"

    Test_RACHelper_Verificar_Cambios_HappyEdicionAllEqual = BuildOk("edicion-equal", logs)
    Exit Function
EH:
    Test_RACHelper_Verificar_Cambios_HappyEdicionAllEqual = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 8. Verificar_Cambios — edicion one field differs -> hayCambios=true
' ---------------------------------------------------------------------------
Public Function Test_RACHelper_Verificar_Cambios_HappyEdicionOneDiff() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("ORIGINAL", "orig@example.com", "edited-desc")
    Dim originales As Object
    Set originales = BuildStubValores("ORIGINAL", "orig@example.com", "original-desc")

    Dim p_Error As String
    Dim json As String
    json = modRACHelper.RAC_Verificar_Cambios(actuales, originales, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hayCambios") <> True Then Err.Raise 1002, , "expected hayCambios=true when one field differs"
    If parsed("payload")("diffs")("DESCRIPCION") <> True Then Err.Raise 1003, , "expected diffs.DESCRIPCION=true"
    If parsed("payload")("diffs")("RAC") <> False Then Err.Raise 1004, , "expected diffs.RAC=false"

    Test_RACHelper_Verificar_Cambios_HappyEdicionOneDiff = BuildOk("edicion-one-diff", logs)
    Exit Function
EH:
    Test_RACHelper_Verificar_Cambios_HappyEdicionOneDiff = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 9. Verificar_Cambios — sad no actuales
' ---------------------------------------------------------------------------
Public Function Test_RACHelper_Verificar_Cambios_SadNoActuales() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modRACHelper.RAC_Verificar_Cambios(Nothing, Nothing, p_Error)
    If p_Error = "" Then Err.Raise 1001, , "expected p_Error populated for Nothing actuales"
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false, got " & json

    Test_RACHelper_Verificar_Cambios_SadNoActuales = BuildOk("no-actuales-rejected", logs)
    Exit Function
EH:
    Test_RACHelper_Verificar_Cambios_SadNoActuales = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 10. Registrar — happy alta (m_ObjRACActivo is Nothing -> insert)
'      Cardinalidad before/after (access-vba-tdd §4.5).
' ---------------------------------------------------------------------------
Public Function Test_RACHelper_Registrar_HappyAlta() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    Set m_ObjRACActivo = Nothing

    Dim altaId As String
    altaId = CStr(TEST_BASE_ID + 1)

    Dim valores As Object
    Set valores = BuildStubValores("ALTA-TEST-RAC", "alta@example.com", "alta-desc")

    Dim p_Error As String
    Dim json As String
    json = modRACHelper.RAC_Registrar(valores, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("ok") <> True Then Err.Raise 1002, , "expected ok=true in payload"
    If CStr(parsed("payload")("modo")) <> "alta" Then Err.Raise 1003, , "expected modo=alta"

    ' Cardinalidad: row exists in TbRACS.
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim countAfter As Long
    countAfter = DCount("*", "TbRACS", "RAC='ALTA-TEST-RAC'")
    If countAfter <> 1 Then Err.Raise 1004, , "expected 1 row after alta, got " & countAfter

    logs(0) = "Registrar_HappyAlta: action=alta, row inserted (count=" & countAfter & ")"

    Test_RACHelper_Registrar_HappyAlta = BuildOk("alta", logs)
    Call TeardownFixture
    Set m_ObjRACActivo = Nothing
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    Set m_ObjRACActivo = Nothing
    On Error GoTo 0
    Test_RACHelper_Registrar_HappyAlta = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 11. Registrar — happy edicion (m_ObjRACActivo set -> update)
' ---------------------------------------------------------------------------
Public Function Test_RACHelper_Registrar_HappyEdicion() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    Dim seedId As String
    seedId = CStr(TEST_BASE_ID + 1)
    Dim insertErr As String
    If Not InsertRealRow(seedId, "ORIGINAL-RAC", "orig@example.com", "orig-desc", insertErr) Then
        Err.Raise 1001, , insertErr
    End If

    ' Set m_ObjRACActivo to the seeded entity so helper knows it's edicion mode.
    Dim active As New RAC
    active.IDRAC = seedId
    Set m_ObjRACActivo = active

    Dim valores As Object
    Set valores = BuildStubValores("EDITED-RAC", "edited@example.com", "edited-desc")

    Dim p_Error As String
    Dim json As String
    json = modRACHelper.RAC_Registrar(valores, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("ok") <> True Then Err.Raise 1002, , "expected ok=true in payload"
    If CStr(parsed("payload")("modo")) <> "edicion" Then Err.Raise 1003, , "expected modo=edicion"

    ' Cardinalidad: row updated to new value.
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim countAfter As Long
    countAfter = DCount("*", "TbRACS", "IDRAC=" & seedId & " AND RAC='EDITED-RAC'")
    If countAfter <> 1 Then Err.Raise 1004, , "expected 1 row with new RAC value, got " & countAfter

    logs(0) = "Registrar_HappyEdicion: action=edicion, row updated (count=" & countAfter & ")"

    Test_RACHelper_Registrar_HappyEdicion = BuildOk("edicion", logs)
    Call TeardownFixture
    Set m_ObjRACActivo = Nothing
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    Set m_ObjRACActivo = Nothing
    On Error GoTo 0
    Test_RACHelper_Registrar_HappyEdicion = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 12. Registrar — sad no valores
' ---------------------------------------------------------------------------
Public Function Test_RACHelper_Registrar_SadNoValores() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Set m_ObjRACActivo = Nothing

    Dim p_Error As String
    Dim json As String
    json = modRACHelper.RAC_Registrar(Nothing, p_Error)
    If p_Error = "" Then Err.Raise 1001, , "expected p_Error populated for Nothing valores"
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false, got " & json

    Test_RACHelper_Registrar_SadNoValores = BuildOk("no-valores-rejected", logs)
    Set m_ObjRACActivo = Nothing
    Exit Function
EH:
    On Error Resume Next
    Set m_ObjRACActivo = Nothing
    On Error GoTo 0
    Test_RACHelper_Registrar_SadNoValores = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 13. Cerrar — happy (m_ObjRACActivo cleared, ok=true)
' ---------------------------------------------------------------------------
Public Function Test_RACHelper_Cerrar_HappyCleared() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    ' Pre-populate m_ObjRACActivo so we can verify it gets cleared.
    Dim preObj As New RAC
    preObj.IDRAC = "999999"
    Set m_ObjRACActivo = preObj

    Dim p_Error As String
    Dim json As String
    json = modRACHelper.RAC_Cerrar(p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    If Not m_ObjRACActivo Is Nothing Then Err.Raise 1002, , "expected m_ObjRACActivo cleared, still set"

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("ok") <> True Then Err.Raise 1003, , "expected payload.ok=true"

    Test_RACHelper_Cerrar_HappyCleared = BuildOk("cleared", logs)
    Exit Function
EH:
    Test_RACHelper_Cerrar_HappyCleared = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 14. RunAll — wrapper for Dysflow manifest discovery
' ---------------------------------------------------------------------------
Public Function Test_RACHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_RACHelper_Abrir_Inicializar_HappyAlta", _
        "Test_RACHelper_Abrir_Inicializar_HappyEdicion", _
        "Test_RACHelper_Abrir_Inicializar_SadInvalidModo", _
        "Test_RACHelper_Abrir_Inicializar_SadEdicionSinID", _
        "Test_RACHelper_Verificar_Cambios_HappyAltaAllEmpty", _
        "Test_RACHelper_Verificar_Cambios_HappyAltaOneField", _
        "Test_RACHelper_Verificar_Cambios_HappyEdicionAllEqual", _
        "Test_RACHelper_Verificar_Cambios_HappyEdicionOneDiff", _
        "Test_RACHelper_Verificar_Cambios_SadNoActuales", _
        "Test_RACHelper_Registrar_HappyAlta", _
        "Test_RACHelper_Registrar_HappyEdicion", _
        "Test_RACHelper_Registrar_SadNoValores", _
        "Test_RACHelper_Cerrar_HappyCleared" _
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
        Test_RACHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_RACHelper_RunAll = BuildOk("all-passed", logs)
    End If
    On Error Resume Next
    Call TeardownFixture
    Set m_ObjRACActivo = Nothing
    On Error GoTo 0
    Exit Function
EH:
    On Error Resume Next
    Dim p_Error As String
    p_Error = "Test_RACHelper_RunAll EH: " & Err.Description
    Call TeardownFixture
    Set m_ObjRACActivo = Nothing
    On Error GoTo 0
    Test_RACHelper_RunAll = BuildFail(p_Error, logs)
End Function
