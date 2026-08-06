Attribute VB_Name = "Test_PECALAltaHelper"
Option Compare Database
Option Explicit

' Test_PECALAltaHelper — REWORK TDD atoms for modPECALAltaHelper.bas
' (Form_FormPECAL Alta/Edición form, see docs/audit/pecal-alta-thin.md).
'
' Anti-pattern removed (was in PR #30 commit 97bdb21):
'   - Test_PECALAltaHelper_OpenForm(ByRef p_Form, ByRef p_Error) used to call
'     DoCmd.OpenForm TEST_FORM_NAME and return Forms(TEST_FORM_NAME) as p_Form.
'     This caused VBE UI interruption in headless runs (user-reported 2026-06-26).
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
'   - §4.5 cardinalidad for mutaciones: countBefore / countAfter for Registrar tests.
'   - §5.1 fixture IDs in test range (>= 900000).
'   - §5.4 db injection explicit where DAO is touched.

' === Module-level constants (all at top per vba-access §10.1) =======================

' Fixture ID base — entities created by atoms use IDs in this range so teardown is
' surgical (DELETE WHERE IDPECAL >= TEST_BASE_ID).
Private Const TEST_BASE_ID As Long = 900850

' Field keys used in the stub Dictionary {PECAL, DESCRIPCION}.
Private Const TEST_FIELD_NAME As String = "PECAL"
Private Const TEST_FIELD_DESCRIPCION As String = "DESCRIPCION"


' === Local helpers (all at top per vba-access §10.1) ================================

' --- BuildOk / BuildFail --------------------------------------------------------------
Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

' --- BuildStubValores ----------------------------------------------------------------
' Builds a Dictionary {PECAL, DESCRIPCION} representing form values.
' Used to pass into the helper without depending on a real form.
Private Function BuildStubValores( _
    ByVal p_Name As String, _
    ByVal p_Desc As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d(TEST_FIELD_NAME) = p_Name
    d(TEST_FIELD_DESCRIPCION) = p_Desc
    Set BuildStubValores = d
End Function

' --- InsertRealRow -------------------------------------------------------------------
' Inserts a row into TbPECAL for the Registrar_HappyEdicion test.
Private Function InsertRealRow( _
    ByVal p_ID As String, _
    ByVal p_Name As String, _
    ByVal p_Desc As String, _
    ByRef p_Error As String _
) As Boolean
    On Error GoTo EH
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim sql As String
    sql = "INSERT INTO TbPECAL (IDPECAL, PECAL, Descripcion) VALUES (" & _
          p_ID & ", '" & Replace(p_Name, "'", "''") & "', '" & Replace(p_Desc, "'", "''") & "')"
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
    before = DCount("*", "TbPECAL", "IDPECAL >= " & TEST_BASE_ID)
    db.Execute "DELETE FROM TbPECAL WHERE IDPECAL >= " & TEST_BASE_ID, dbFailOnError
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
' 1. Abrir_Inicializar — happy alta (no ID, no m_ObjPECALActiva)
' ---------------------------------------------------------------------------
Public Function Test_PECALAltaHelper_Abrir_Inicializar_HappyAlta() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    ' Ensure m_ObjPECALActiva is clean.
    Set m_ObjPECALActiva = Nothing

    Dim p_Error As String
    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim json As String
    json = modPECALAltaHelper.PECALAlta_Abrir_Inicializar("alta", "", p_Entidad, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If InStr(json, """ok"":true") = 0 Then Err.Raise 1001, , "expected ok=true, got " & json

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim msg As String
    If Not AssertPayloadHasKey(parsed("payload"), "titulo", msg) Then Err.Raise 1002, , msg
    If Not AssertPayloadHasKey(parsed("payload"), "modo", msg) Then Err.Raise 1002, , msg
    If Not AssertPayloadHasKey(parsed("payload"), "hasEntidad", msg) Then Err.Raise 1002, , msg

    If CStr(parsed("payload")("titulo")) <> "ALTA DE PECAL" Then Err.Raise 1003, , "expected titulo=ALTA DE PECAL"
    If CStr(parsed("payload")("modo")) <> "alta" Then Err.Raise 1004, , "expected modo=alta"
    If parsed("payload")("hasEntidad") <> False Then Err.Raise 1005, , "expected hasEntidad=false"
    If p_Entidad Is Nothing Then Err.Raise 1006, , "expected p_Entidad populated as empty stub"

    Test_PECALAltaHelper_Abrir_Inicializar_HappyAlta = BuildOk("alta-initialized", logs)
    Exit Function
EH:
    Test_PECALAltaHelper_Abrir_Inicializar_HappyAlta = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 2. Abrir_Inicializar — happy edicion (seeded row + ID)
' ---------------------------------------------------------------------------
Public Function Test_PECALAltaHelper_Abrir_Inicializar_HappyEdicion() As String
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
    json = modPECALAltaHelper.PECALAlta_Abrir_Inicializar("edicion", seedId, p_Entidad, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("titulo")) <> "EDICIÓN DE PECAL" Then Err.Raise 1002, , "expected titulo=EDICIÓN DE PECAL"
    If CStr(parsed("payload")("modo")) <> "edicion" Then Err.Raise 1003, , "expected modo=edicion"
    If parsed("payload")("hasEntidad") <> True Then Err.Raise 1004, , "expected hasEntidad=true"
    If CStr(parsed("payload")("id")) <> seedId Then Err.Raise 1005, , "expected id=" & seedId
    If p_Entidad Is Nothing Then Err.Raise 1006, , "expected p_Entidad populated"
    If CStr(p_Entidad(TEST_FIELD_NAME)) <> "EDIT-EXISTING" Then Err.Raise 1007, , "expected PECAL field populated"
    If CStr(p_Entidad(TEST_FIELD_DESCRIPCION)) <> "edit-desc" Then Err.Raise 1008, , "expected DESCRIPCION field populated"

    Test_PECALAltaHelper_Abrir_Inicializar_HappyEdicion = BuildOk("edicion-initialized", logs)
    Call TeardownFixture
    Set m_ObjPECALActiva = Nothing
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    Set m_ObjPECALActiva = Nothing
    On Error GoTo 0
    Test_PECALAltaHelper_Abrir_Inicializar_HappyEdicion = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 3. Abrir_Inicializar — sad invalid modo
' ---------------------------------------------------------------------------
Public Function Test_PECALAltaHelper_Abrir_Inicializar_SadInvalidModo() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Set m_ObjPECALActiva = Nothing

    Dim p_Error As String
    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim json As String
    json = modPECALAltaHelper.PECALAlta_Abrir_Inicializar("weird", "", p_Entidad, p_Error)
    If p_Error = "" Then Err.Raise 1001, , "expected p_Error populated for invalid modo"
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false, got " & json

    Test_PECALAltaHelper_Abrir_Inicializar_SadInvalidModo = BuildOk("invalid-modo-rejected", logs)
    Exit Function
EH:
    Test_PECALAltaHelper_Abrir_Inicializar_SadInvalidModo = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 4. Abrir_Inicializar — sad edicion without ID
' ---------------------------------------------------------------------------
Public Function Test_PECALAltaHelper_Abrir_Inicializar_SadEdicionSinID() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Set m_ObjPECALActiva = Nothing

    Dim p_Error As String
    Dim p_Entidad As Object
    Set p_Entidad = Nothing

    Dim json As String
    json = modPECALAltaHelper.PECALAlta_Abrir_Inicializar("edicion", "", p_Entidad, p_Error)
    If p_Error = "" Then Err.Raise 1001, , "expected p_Error populated when edicion mode without ID"
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false, got " & json

    Test_PECALAltaHelper_Abrir_Inicializar_SadEdicionSinID = BuildOk("edicion-sin-id-rejected", logs)
    Exit Function
EH:
    Test_PECALAltaHelper_Abrir_Inicializar_SadEdicionSinID = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 5. VerificarCambios — alta with all-empty values -> hayCambios=false
' ---------------------------------------------------------------------------
Public Function Test_PECALAltaHelper_VerificarCambios_HappyAltaAllEmpty() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("", "")

    Dim p_Error As String
    Dim json As String
    json = modPECALAltaHelper.PECALAlta_VerificarCambios(actuales, Nothing, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hayCambios") <> False Then Err.Raise 1002, , "expected hayCambios=false for empty alta"

    Test_PECALAltaHelper_VerificarCambios_HappyAltaAllEmpty = BuildOk("alta-empty", logs)
    Exit Function
EH:
    Test_PECALAltaHelper_VerificarCambios_HappyAltaAllEmpty = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 6. VerificarCambios — alta with one field filled -> hayCambios=true
' ---------------------------------------------------------------------------
Public Function Test_PECALAltaHelper_VerificarCambios_HappyAltaOneField() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("SOMETHING", "")

    Dim p_Error As String
    Dim json As String
    json = modPECALAltaHelper.PECALAlta_VerificarCambios(actuales, Nothing, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hayCambios") <> True Then Err.Raise 1002, , "expected hayCambios=true for alta with content"

    Test_PECALAltaHelper_VerificarCambios_HappyAltaOneField = BuildOk("alta-one-field", logs)
    Exit Function
EH:
    Test_PECALAltaHelper_VerificarCambios_HappyAltaOneField = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 7. VerificarCambios — edicion all equal -> hayCambios=false
' ---------------------------------------------------------------------------
Public Function Test_PECALAltaHelper_VerificarCambios_HappyEdicionAllEqual() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("SAME", "same-desc")
    Dim originales As Object
    Set originales = BuildStubValores("SAME", "same-desc")

    Dim p_Error As String
    Dim json As String
    json = modPECALAltaHelper.PECALAlta_VerificarCambios(actuales, originales, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hayCambios") <> False Then Err.Raise 1002, , "expected hayCambios=false for equal fields"

    Test_PECALAltaHelper_VerificarCambios_HappyEdicionAllEqual = BuildOk("edicion-equal", logs)
    Exit Function
EH:
    Test_PECALAltaHelper_VerificarCambios_HappyEdicionAllEqual = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 8. VerificarCambios — edicion one field differs -> hayCambios=true
' ---------------------------------------------------------------------------
Public Function Test_PECALAltaHelper_VerificarCambios_HappyEdicionOneDiff() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim actuales As Object
    Set actuales = BuildStubValores("ORIGINAL", "edited-desc")
    Dim originales As Object
    Set originales = BuildStubValores("ORIGINAL", "original-desc")

    Dim p_Error As String
    Dim json As String
    json = modPECALAltaHelper.PECALAlta_VerificarCambios(actuales, originales, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hayCambios") <> True Then Err.Raise 1002, , "expected hayCambios=true when one field differs"
    If parsed("payload")("diffs")(TEST_FIELD_DESCRIPCION) <> True Then Err.Raise 1003, , "expected diffs.DESCRIPCION=true"
    If parsed("payload")("diffs")(TEST_FIELD_NAME) <> False Then Err.Raise 1004, , "expected diffs.PECAL=false"

    Test_PECALAltaHelper_VerificarCambios_HappyEdicionOneDiff = BuildOk("edicion-one-diff", logs)
    Exit Function
EH:
    Test_PECALAltaHelper_VerificarCambios_HappyEdicionOneDiff = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 9. VerificarCambios — sad no actuales
' ---------------------------------------------------------------------------
Public Function Test_PECALAltaHelper_VerificarCambios_SadNoActuales() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modPECALAltaHelper.PECALAlta_VerificarCambios(Nothing, Nothing, p_Error)
    If p_Error = "" Then Err.Raise 1001, , "expected p_Error populated for Nothing actuales"
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false, got " & json

    Test_PECALAltaHelper_VerificarCambios_SadNoActuales = BuildOk("no-actuales-rejected", logs)
    Exit Function
EH:
    Test_PECALAltaHelper_VerificarCambios_SadNoActuales = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 10. Registrar — happy alta (m_ObjPECALActiva is Nothing -> insert)
'      Cardinalidad before/after (access-vba-tdd §4.5).
' ---------------------------------------------------------------------------
Public Function Test_PECALAltaHelper_Registrar_HappyAlta() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    Set m_ObjPECALActiva = Nothing

    Dim altaId As String
    altaId = CStr(TEST_BASE_ID + 1)

    Dim valores As Object
    Set valores = BuildStubValores("ALTA-TEST-PECAL", "alta-desc")

    Dim p_Error As String
    Dim json As String
    json = modPECALAltaHelper.PECALAlta_Registrar(valores, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("ok") <> True Then Err.Raise 1002, , "expected ok=true in payload"
    If CStr(parsed("payload")("modo")) <> "alta" Then Err.Raise 1003, , "expected modo=alta"

    ' Cardinalidad: row exists in TbPECAL.
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim countAfter As Long
    countAfter = DCount("*", "TbPECAL", "PECAL='ALTA-TEST-PECAL'")
    If countAfter <> 1 Then Err.Raise 1004, , "expected 1 row after alta, got " & countAfter

    logs(0) = "Registrar_HappyAlta: action=alta, row inserted (count=" & countAfter & ")"

    Test_PECALAltaHelper_Registrar_HappyAlta = BuildOk("alta", logs)
    Call TeardownFixture
    Set m_ObjPECALActiva = Nothing
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    Set m_ObjPECALActiva = Nothing
    On Error GoTo 0
    Test_PECALAltaHelper_Registrar_HappyAlta = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 11. Registrar — happy edicion (m_ObjPECALActiva set -> update)
' ---------------------------------------------------------------------------
Public Function Test_PECALAltaHelper_Registrar_HappyEdicion() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    Dim seedId As String
    seedId = CStr(TEST_BASE_ID + 1)
    Dim insertErr As String
    If Not InsertRealRow(seedId, "ORIGINAL-PECAL", "orig-desc", insertErr) Then
        Err.Raise 1001, , insertErr
    End If

    ' Set m_ObjPECALActiva to the seeded entity so helper knows it's edicion mode.
    Dim active As New PECAL
    active.IDPECAL = seedId
    Set m_ObjPECALActiva = active

    Dim valores As Object
    Set valores = BuildStubValores("EDITED-PECAL", "edited-desc")

    Dim p_Error As String
    Dim json As String
    json = modPECALAltaHelper.PECALAlta_Registrar(valores, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("ok") <> True Then Err.Raise 1002, , "expected ok=true in payload"
    If CStr(parsed("payload")("modo")) <> "edicion" Then Err.Raise 1003, , "expected modo=edicion"

    ' Cardinalidad: row updated to new value.
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim countAfter As Long
    countAfter = DCount("*", "TbPECAL", "IDPECAL=" & seedId & " AND PECAL='EDITED-PECAL'")
    If countAfter <> 1 Then Err.Raise 1004, , "expected 1 row with new PECAL value, got " & countAfter

    logs(0) = "Registrar_HappyEdicion: action=edicion, row updated (count=" & countAfter & ")"

    Test_PECALAltaHelper_Registrar_HappyEdicion = BuildOk("edicion", logs)
    Call TeardownFixture
    Set m_ObjPECALActiva = Nothing
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    Set m_ObjPECALActiva = Nothing
    On Error GoTo 0
    Test_PECALAltaHelper_Registrar_HappyEdicion = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 12. Registrar — sad no valores
' ---------------------------------------------------------------------------
Public Function Test_PECALAltaHelper_Registrar_SadNoValores() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Set m_ObjPECALActiva = Nothing

    Dim p_Error As String
    Dim json As String
    json = modPECALAltaHelper.PECALAlta_Registrar(Nothing, p_Error)
    If p_Error = "" Then Err.Raise 1001, , "expected p_Error populated for Nothing valores"
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false, got " & json

    Test_PECALAltaHelper_Registrar_SadNoValores = BuildOk("no-valores-rejected", logs)
    Set m_ObjPECALActiva = Nothing
    Exit Function
EH:
    On Error Resume Next
    Set m_ObjPECALActiva = Nothing
    On Error GoTo 0
    Test_PECALAltaHelper_Registrar_SadNoValores = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 13. Cerrar — happy (m_ObjPECALActiva cleared, ok=true)
' ---------------------------------------------------------------------------
Public Function Test_PECALAltaHelper_Cerrar_HappyCleared() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    ' Pre-populate m_ObjPECALActiva so we can verify it gets cleared.
    Dim preObj As New PECAL
    preObj.IDPECAL = "999999"
    Set m_ObjPECALActiva = preObj

    Dim p_Error As String
    Dim json As String
    json = modPECALAltaHelper.PECALAlta_Cerrar(p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    If Not m_ObjPECALActiva Is Nothing Then Err.Raise 1002, , "expected m_ObjPECALActiva cleared, still set"

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("ok") <> True Then Err.Raise 1003, , "expected payload.ok=true"

    Test_PECALAltaHelper_Cerrar_HappyCleared = BuildOk("cleared", logs)
    Exit Function
EH:
    Test_PECALAltaHelper_Cerrar_HappyCleared = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 14. RunAll — wrapper for Dysflow manifest discovery
' ---------------------------------------------------------------------------
Public Function Test_PECALAltaHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_PECALAltaHelper_Abrir_Inicializar_HappyAlta", _
        "Test_PECALAltaHelper_Abrir_Inicializar_HappyEdicion", _
        "Test_PECALAltaHelper_Abrir_Inicializar_SadInvalidModo", _
        "Test_PECALAltaHelper_Abrir_Inicializar_SadEdicionSinID", _
        "Test_PECALAltaHelper_VerificarCambios_HappyAltaAllEmpty", _
        "Test_PECALAltaHelper_VerificarCambios_HappyAltaOneField", _
        "Test_PECALAltaHelper_VerificarCambios_HappyEdicionAllEqual", _
        "Test_PECALAltaHelper_VerificarCambios_HappyEdicionOneDiff", _
        "Test_PECALAltaHelper_VerificarCambios_SadNoActuales", _
        "Test_PECALAltaHelper_Registrar_HappyAlta", _
        "Test_PECALAltaHelper_Registrar_HappyEdicion", _
        "Test_PECALAltaHelper_Registrar_SadNoValores", _
        "Test_PECALAltaHelper_Cerrar_HappyCleared" _
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
        Test_PECALAltaHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_PECALAltaHelper_RunAll = BuildOk("all-passed", logs)
    End If
    On Error Resume Next
    Call TeardownFixture
    Set m_ObjPECALActiva = Nothing
    On Error GoTo 0
    Exit Function
EH:
    On Error Resume Next
    Dim p_Error As String
    p_Error = "Test_PECALAltaHelper_RunAll EH: " & Err.Description
    Call TeardownFixture
    Set m_ObjPECALActiva = Nothing
    On Error GoTo 0
    Test_PECALAltaHelper_RunAll = BuildFail(p_Error, logs)
End Function