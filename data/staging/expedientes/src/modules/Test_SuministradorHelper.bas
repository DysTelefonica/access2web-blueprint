Attribute VB_Name = "Test_SuministradorHelper"
Option Compare Database
Option Explicit

' Test_SuministradorHelper — REWORK (2026-06-26)
' Pure-data TDD atoms for modSuministradorHelper.bas
' (Form_FormSuministradoresGestion, see docs/audit/suministradores-gestion-thin.md).
'
' Anti-pattern removed (was in PR #31 commit c86d460):
'   - Test_SuministradorHelper_OpenForm(ByRef p_Form, ByRef p_Error) used to call
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
'   - §1.8 declaration ordering: all Private Const/Function/Sub at top, Public atoms after.
'   - §1.10 atom signature MUST match helper signature EXACTLY (audited in §1.9 pre-compile).
'   - §4.2 no-humo: atoms assert concrete values (counts, fields, JSON), not "did not crash".
'   - §4.4 strong assertions: count=0 vs >0, payload keys present, etc.
'   - §4.5 cardinalidad for mutaciones: countBefore / countAfter for Eliminar tests.
'   - §5.1 fixture IDs in test range (>= 900000).
'   - §5.4 db injection explicit where DAO is touched.
'
' Convention:
'   - Atoms use ONLY Global public names (access-vba-tdd §1.1.1).
'   - Each atom returns JSON: {"ok":true|false,"value":...,"payload":null,"error":...,"logs":[...]}.
'   - Atoms never call MsgBox or pop up UI.

' === Module-level constants (all at top per vba-access §10.1) =======================

' Fixture ID base — entities created by atoms use IDs in this range so teardown is
' surgical (DELETE WHERE IDSuministrador >= TEST_BASE_ID).
Private Const TEST_BASE_ID As Long = 900700

' Field separator used in rowSource assertions. Mirrors the helper's constant.
Private Const TEST_FIELD_SEP As String = ";"


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

' --- BuildStubSuministrador ----------------------------------------------------------
' Builds a Scripting.Dictionary that mimics the shape of a Suministrador entity.
' Returns a Dictionary {IDSuministrador, CIF, Nombre}.
Private Function BuildStubSuministrador( _
    ByVal p_ID As String, _
    ByVal p_CIF As String, _
    ByVal p_Nombre As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("IDSuministrador") = p_ID
    d("CIF") = p_CIF
    d("Nombre") = p_Nombre
    Set BuildStubSuministrador = d
End Function

' --- BuildStubSuministradoresDict ----------------------------------------------------
' Builds a Dictionary {CStr(ID) -> stubEntity} from an array of {CIF, Nombre} tuples.
' IDs are TEST_BASE_ID, TEST_BASE_ID+1, ... so each row has a unique key.
Private Function BuildStubSuministradoresDict(ByVal p_Items As Variant) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(p_Items) To UBound(p_Items)
        Dim idStr As String
        idStr = CStr(TEST_BASE_ID + i)
        Set d(idStr) = BuildStubSuministrador(idStr, _
                                              CStr(p_Items(i)(0)), _
                                              CStr(p_Items(i)(1)))
    Next i
    Set BuildStubSuministradoresDict = d
End Function

' --- InsertRealRow -------------------------------------------------------------------
' Inserts a row into TbSuministradores for the Eliminar_Borrar happy/adversarial tests
' (which exercise the DAO path via Helper_EntidadCRUD). Returns True on success.
Private Function InsertRealRow( _
    ByVal p_ID As String, _
    ByVal p_Nombre As String, _
    ByVal p_CIF As String, _
    ByRef p_Error As String _
) As Boolean
    On Error GoTo EH
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim sql As String
    sql = "INSERT INTO TbSuministradores (IDSuministrador, Nombre, CIF) VALUES (" & _
          p_ID & ", '" & Replace(p_Nombre, "'", "''") & "', '" & Replace(p_CIF, "'", "''") & "')"
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
    before = DCount("*", "TbSuministradores", "IDSuministrador >= " & TEST_BASE_ID)
    db.Execute "DELETE FROM TbSuministradores WHERE IDSuministrador >= " & TEST_BASE_ID, dbFailOnError
    On Error GoTo 0
    TeardownFixture = before
End Function

' --- AssertPayloadHasKey --------------------------------------------------------------
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
' 1. Abrir_Inicializar — happy (admin=true, no OpenArgs)
'    Expected: showAlta=true, focusAlta=true, showElegir=false
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Abrir_Inicializar_HappyAdminNoArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Abrir_Inicializar( _
        True, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If InStr(json, """ok"":true") = 0 Then Err.Raise 1001, , "expected ok=true, got " & json

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim msg As String
    If Not AssertPayloadHasKey(parsed("payload"), "showAlta", msg) Then Err.Raise 1002, , msg
    If Not AssertPayloadHasKey(parsed("payload"), "focusAlta", msg) Then Err.Raise 1002, , msg
    If Not AssertPayloadHasKey(parsed("payload"), "showElegir", msg) Then Err.Raise 1002, , msg

    If parsed("payload")("showAlta") <> True Then Err.Raise 1003, , "expected showAlta=true"
    If parsed("payload")("focusAlta") <> True Then Err.Raise 1003, , "expected focusAlta=true"
    If parsed("payload")("showElegir") <> False Then Err.Raise 1003, , "expected showElegir=false"

    Test_SuministradorHelper_Abrir_Inicializar_HappyAdminNoArgs = BuildOk("admin-no-args", logs)
    Exit Function
EH:
    Test_SuministradorHelper_Abrir_Inicializar_HappyAdminNoArgs = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 2. Abrir_Inicializar — happy (admin=false, no OpenArgs)
'    Expected: showAlta=false, focusAlta=false, showElegir=false
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Abrir_Inicializar_HappyNonAdminNoArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Abrir_Inicializar( _
        False, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("showAlta") <> False Then Err.Raise 1002, , "expected showAlta=false"
    If parsed("payload")("focusAlta") <> False Then Err.Raise 1002, , "expected focusAlta=false"
    If parsed("payload")("showElegir") <> False Then Err.Raise 1002, , "expected showElegir=false"

    Test_SuministradorHelper_Abrir_Inicializar_HappyNonAdminNoArgs = BuildOk("non-admin-no-args", logs)
    Exit Function
EH:
    Test_SuministradorHelper_Abrir_Inicializar_HappyNonAdminNoArgs = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 3. Abrir_Inicializar — happy (admin=true, has OpenArgs)
'    Expected: showAlta=true, focusAlta=true, showElegir=true
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Abrir_Inicializar_HappyAdminWithArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Abrir_Inicializar( _
        True, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("showAlta") <> True Then Err.Raise 1002, , "expected showAlta=true"
    If parsed("payload")("focusAlta") <> True Then Err.Raise 1002, , "expected focusAlta=true"
    If parsed("payload")("showElegir") <> True Then Err.Raise 1002, , "expected showElegir=true"

    Test_SuministradorHelper_Abrir_Inicializar_HappyAdminWithArgs = BuildOk("admin-with-args", logs)
    Exit Function
EH:
    Test_SuministradorHelper_Abrir_Inicializar_HappyAdminWithArgs = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 4. Buscar_Listar — happy (3 entities, no filter -> 3 rows)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Buscar_Listar_HappyAllRows() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    ' (CIF, Nombre) tuples
    Set col = BuildStubSuministradoresDict(Array( _
        Array("A12345678", "ALPHA-SUM"), _
        Array("B12345678", "BETA-SUM"), _
        Array("G12345678", "GAMMA-SUM")))

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Buscar_Listar(col, "", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim msg As String
    If Not AssertPayloadHasKey(parsed("payload"), "count", msg) Then Err.Raise 1001, , msg
    If Not AssertPayloadHasKey(parsed("payload"), "rowSource", msg) Then Err.Raise 1001, , msg

    If CLng(parsed("payload")("count")) <> 3 Then Err.Raise 1002, , "expected count=3, got " & parsed("payload")("count")

    Dim rs As String
    rs = CStr(parsed("payload")("rowSource"))
    If InStr(rs, "ALPHA-SUM") = 0 Then Err.Raise 1003, , "rowSource missing ALPHA-SUM"
    If InStr(rs, "BETA-SUM") = 0 Then Err.Raise 1003, , "rowSource missing BETA-SUM"
    If InStr(rs, "GAMMA-SUM") = 0 Then Err.Raise 1003, , "rowSource missing GAMMA-SUM"

    Test_SuministradorHelper_Buscar_Listar_HappyAllRows = BuildOk(3, logs)
    Exit Function
EH:
    Test_SuministradorHelper_Buscar_Listar_HappyAllRows = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 5. Buscar_Listar — edge (empty collection -> 0 rows)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Buscar_Listar_EdgeEmptyCollection() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Buscar_Listar(Nothing, "", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 0 Then Err.Raise 1002, , "expected count=0"

    Test_SuministradorHelper_Buscar_Listar_EdgeEmptyCollection = BuildOk(0, logs)
    Exit Function
EH:
    Test_SuministradorHelper_Buscar_Listar_EdgeEmptyCollection = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 6. Buscar_Listar — happy (filter matches by Nombre substring)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Buscar_Listar_HappyFilterByNombre() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubSuministradoresDict(Array( _
        Array("X11111111", "ALPHA-A"), _
        Array("Y22222222", "BETA"), _
        Array("Z33333333", "ALPHA-B")))

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Buscar_Listar(col, "ALPHA", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 2 Then Err.Raise 1002, , "expected count=2 (ALPHA Nombre match), got " & parsed("payload")("count")

    Dim rs As String
    rs = CStr(parsed("payload")("rowSource"))
    If InStr(rs, "BETA") > 0 Then Err.Raise 1003, , "rowSource should NOT contain BETA"

    Test_SuministradorHelper_Buscar_Listar_HappyFilterByNombre = BuildOk(2, logs)
    Exit Function
EH:
    Test_SuministradorHelper_Buscar_Listar_HappyFilterByNombre = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 7. Buscar_Listar — happy (filter matches by CIF substring, dual-field)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Buscar_Listar_HappyFilterByCIF() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubSuministradoresDict(Array( _
        Array("B11111111", "ALPHA-NAME"), _
        Array("X99999999", "BETA-NAME"), _
        Array("B22222222", "GAMMA-NAME")))

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Buscar_Listar(col, "B", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    ' Two CIFs start with B; BETA-NAME does not (CIF starts with X).
    If CLng(parsed("payload")("count")) <> 2 Then
        Err.Raise 1002, , "expected count=2 (B CIF prefix), got " & parsed("payload")("count")
    End If

    Dim rs As String
    rs = CStr(parsed("payload")("rowSource"))
    If InStr(rs, "BETA-NAME") > 0 Then
        Err.Raise 1003, , "rowSource should NOT contain BETA-NAME (CIF X99999999)"
    End If

    Test_SuministradorHelper_Buscar_Listar_HappyFilterByCIF = BuildOk(2, logs)
    Exit Function
EH:
    Test_SuministradorHelper_Buscar_Listar_HappyFilterByCIF = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 8. Buscar_Listar — edge (256-char filter that does not match anything -> 0 rows)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Buscar_Listar_EdgeLongFilter() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubSuministradoresDict(Array( _
        Array("A11111111", "SHORT-NAME")))

    Dim longFilter As String
    longFilter = String(256, "X")

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Buscar_Listar(col, longFilter, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 0 Then Err.Raise 1002, , "expected count=0 for 256-char non-matching filter"

    Test_SuministradorHelper_Buscar_Listar_EdgeLongFilter = BuildOk(0, logs)
    Exit Function
EH:
    Test_SuministradorHelper_Buscar_Listar_EdgeLongFilter = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 9. Buscar_Listar — adversarial (semicolons in CIF are sanitized to colons)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Buscar_Listar_AdversarialSemicolonInCIF() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = CreateObject("Scripting.Dictionary")
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("IDSuministrador") = CStr(TEST_BASE_ID + 1)
    d("CIF") = "A;B;C"
    d("Nombre") = "evil;name"
    Set col(CStr(TEST_BASE_ID + 1)) = d

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Buscar_Listar(col, "", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    Dim rs As String
    rs = CStr(parsed("payload")("rowSource"))
    If InStr(rs, "A;B;C") > 0 Then Err.Raise 1002, , "expected semicolons sanitized in CIF"
    If InStr(rs, "A:B:C") = 0 Then Err.Raise 1003, , "expected sanitized value A:B:C"

    Test_SuministradorHelper_Buscar_Listar_AdversarialSemicolonInCIF = BuildOk(1, logs)
    Exit Function
EH:
    Test_SuministradorHelper_Buscar_Listar_AdversarialSemicolonInCIF = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 10. Seleccionar_Cargar — happy (selection present, admin)
'      Expected: entity=stub, enableEditar=true, enableEliminar=true
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Seleccionar_Cargar_HappyAdmin() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubSuministradoresDict(Array( _
        Array("A11111111", "FIRST")))

    Dim selectedId As String
    selectedId = CStr(TEST_BASE_ID + 1)

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Seleccionar_Cargar( _
        selectedId, col, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim msg As String
    If Not AssertPayloadHasKey(parsed("payload"), "enableEditar", msg) Then Err.Raise 1001, , msg
    If Not AssertPayloadHasKey(parsed("payload"), "enableEliminar", msg) Then Err.Raise 1001, , msg
    If Not AssertPayloadHasKey(parsed("payload"), "entity", msg) Then Err.Raise 1001, , msg

    If parsed("payload")("enableEditar") <> True Then Err.Raise 1002, , "expected enableEditar=true"
    If parsed("payload")("enableEliminar") <> True Then Err.Raise 1003, , "expected enableEliminar=true for admin"
    If parsed("payload")("entity") Is Nothing Then Err.Raise 1004, , "expected entity populated"
    If CStr(parsed("payload")("entity")("IDSuministrador")) <> selectedId Then
        Err.Raise 1005, , "expected ID=" & selectedId & ", got " & parsed("payload")("entity")("IDSuministrador")
    End If
    If CStr(parsed("payload")("entity")("Nombre")) <> "FIRST" Then
        Err.Raise 1006, , "expected Nombre=FIRST, got " & parsed("payload")("entity")("Nombre")
    End If

    Test_SuministradorHelper_Seleccionar_Cargar_HappyAdmin = BuildOk(selectedId, logs)
    Exit Function
EH:
    Test_SuministradorHelper_Seleccionar_Cargar_HappyAdmin = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 11. Seleccionar_Cargar — happy (selection present, non-admin)
'      Expected: enableEditar=true, enableEliminar=false
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Seleccionar_Cargar_HappyNonAdmin() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubSuministradoresDict(Array( _
        Array("A11111111", "FIRST")))

    Dim selectedId As String
    selectedId = CStr(TEST_BASE_ID + 1)

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Seleccionar_Cargar( _
        selectedId, col, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("enableEditar") <> True Then Err.Raise 1002, , "expected enableEditar=true"
    If parsed("payload")("enableEliminar") <> False Then Err.Raise 1003, , "expected enableEliminar=false for non-admin"

    Test_SuministradorHelper_Seleccionar_Cargar_HappyNonAdmin = BuildOk(selectedId, logs)
    Exit Function
EH:
    Test_SuministradorHelper_Seleccionar_Cargar_HappyNonAdmin = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 12. Seleccionar_Cargar — sad (empty selection -> both disabled, entity=null)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Seleccionar_Cargar_SadEmptySelection() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubSuministradoresDict(Array( _
        Array("A11111111", "FIRST")))

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Seleccionar_Cargar( _
        "", col, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("enableEditar") <> False Then Err.Raise 1002, , "expected enableEditar=false"
    If parsed("payload")("enableEliminar") <> False Then Err.Raise 1003, , "expected enableEliminar=false"
    If Not parsed("payload")("entity") Is Nothing Then Err.Raise 1004, , "expected entity=null"

    Test_SuministradorHelper_Seleccionar_Cargar_SadEmptySelection = BuildOk("empty", logs)
    Exit Function
EH:
    Test_SuministradorHelper_Seleccionar_Cargar_SadEmptySelection = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 13. Seleccionar_Cargar — sad (selected ID not found -> entity=null, both disabled)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Seleccionar_Cargar_SadNotFound() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubSuministradoresDict(Array( _
        Array("A11111111", "FIRST")))

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Seleccionar_Cargar( _
        "999999", col, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("enableEditar") <> False Then Err.Raise 1002, , "expected enableEditar=false"
    If Not parsed("payload")("entity") Is Nothing Then Err.Raise 1004, , "expected entity=null"

    Test_SuministradorHelper_Seleccionar_Cargar_SadNotFound = BuildOk("not-found", logs)
    Exit Function
EH:
    Test_SuministradorHelper_Seleccionar_Cargar_SadNotFound = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 14. Eliminar_Borrar — sad (p_PromptResult=vbNo -> cancelled, row preserved)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Eliminar_Borrar_SadCancelled() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    On Error GoTo EH

    Dim seedId As String
    seedId = CStr(TEST_BASE_ID + 1)
    Dim insertErr As String
    If Not InsertRealRow(seedId, "CANCEL-TEST", "CANC-CIF", insertErr) Then
        Err.Raise 1001, , insertErr
    End If

    Dim entity As Object
    Set entity = BuildStubSuministrador(seedId, "CANCEL-TEST", "CANC-CIF")

    Dim promptResult As Long
    promptResult = vbNo

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Eliminar_Borrar(entity, promptResult, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    ' Helper returns ok=true with no payload for "cancelled" path
    If InStr(json, """ok"":true") = 0 Then Err.Raise 1002, , "expected ok=true for cancelled path, got " & json

    ' Cardinalidad: row must still exist.
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim countAfter As Long
    countAfter = DCount("*", "TbSuministradores", "IDSuministrador=" & seedId)
    If countAfter <> 1 Then Err.Raise 1003, , "row should still exist after cancel, count=" & countAfter

    Test_SuministradorHelper_Eliminar_Borrar_SadCancelled = BuildOk("cancelled", logs)
    Call TeardownFixture
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    On Error GoTo 0
    Test_SuministradorHelper_Eliminar_Borrar_SadCancelled = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 15. Eliminar_Borrar — happy (p_PromptResult=vbYes -> deleted, row gone)
'      Cardinalidad before/after (access-vba-tdd §4.5).
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Eliminar_Borrar_HappyDeleted() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    Dim seedId As String
    seedId = CStr(TEST_BASE_ID + 1)
    Dim insertErr As String
    If Not InsertRealRow(seedId, "DELETE-TEST", "DEL-CIF", insertErr) Then
        Err.Raise 1001, , insertErr
    End If

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim countBefore As Long
    countBefore = DCount("*", "TbSuministradores", "IDSuministrador=" & seedId)
    If countBefore <> 1 Then Err.Raise 1002, , "expected 1 row before delete, got " & countBefore

    Dim entity As Object
    Set entity = BuildStubSuministrador(seedId, "DELETE-TEST", "DEL-CIF")

    Dim promptResult As Long
    promptResult = vbYes

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Eliminar_Borrar(entity, promptResult, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If InStr(json, """ok"":true") = 0 Then Err.Raise 1003, , "expected ok=true for deleted path, got " & json

    Dim countAfter As Long
    countAfter = DCount("*", "TbSuministradores", "IDSuministrador=" & seedId)
    If countAfter <> 0 Then Err.Raise 1004, , "expected 0 rows after delete, got " & countAfter

    logs(0) = "Eliminar_Borrar_HappyDeleted: cardinality " & countBefore & "->" & countAfter

    Test_SuministradorHelper_Eliminar_Borrar_HappyDeleted = BuildOk("deleted", logs)
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    On Error GoTo 0
    Test_SuministradorHelper_Eliminar_Borrar_HappyDeleted = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 16. Eliminar_Borrar — sad (no entity supplied -> fail JSON)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Eliminar_Borrar_SadNoEntity() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Eliminar_Borrar(Nothing, vbYes, p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false, got " & json

    Test_SuministradorHelper_Eliminar_Borrar_SadNoEntity = BuildOk("no-entity-rejected", logs)
    Exit Function
EH:
    Test_SuministradorHelper_Eliminar_Borrar_SadNoEntity = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 17. Eliminar_Borrar — sad (missing ID -> fail JSON)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_Eliminar_Borrar_SadMissingID() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim entity As Object
    Set entity = BuildStubSuministrador("", "MISSING-ID", "X")

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_Eliminar_Borrar(entity, vbYes, p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false for missing ID, got " & json

    Test_SuministradorHelper_Eliminar_Borrar_SadMissingID = BuildOk("missing-id-rejected", logs)
    Exit Function
EH:
    Test_SuministradorHelper_Eliminar_Borrar_SadMissingID = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 18. DobleClick_AbrirEdicion — happy (choose dispatched when cmdElegir visible)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_DobleClick_AbrirEdicion_HappyChoose() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_DobleClick_AbrirEdicion( _
        True, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("action")) <> "choose" Then
        Err.Raise 1002, , "expected action=choose, got " & parsed("payload")("action")
    End If

    Test_SuministradorHelper_DobleClick_AbrirEdicion_HappyChoose = BuildOk("choose", logs)
    Exit Function
EH:
    Test_SuministradorHelper_DobleClick_AbrirEdicion_HappyChoose = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 19. DobleClick_AbrirEdicion — happy (edit dispatched when ComandoEditar enabled)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_DobleClick_AbrirEdicion_HappyEdit() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_DobleClick_AbrirEdicion( _
        False, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("action")) <> "edit" Then
        Err.Raise 1002, , "expected action=edit, got " & parsed("payload")("action")
    End If

    Test_SuministradorHelper_DobleClick_AbrirEdicion_HappyEdit = BuildOk("edit", logs)
    Exit Function
EH:
    Test_SuministradorHelper_DobleClick_AbrirEdicion_HappyEdit = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 20. DobleClick_AbrirEdicion — happy (none when both buttons unavailable)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_DobleClick_AbrirEdicion_HappyNone() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_DobleClick_AbrirEdicion( _
        False, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("action")) <> "none" Then
        Err.Raise 1002, , "expected action=none, got " & parsed("payload")("action")
    End If

    Test_SuministradorHelper_DobleClick_AbrirEdicion_HappyNone = BuildOk("none", logs)
    Exit Function
EH:
    Test_SuministradorHelper_DobleClick_AbrirEdicion_HappyNone = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 21. DobleClick_AbrirEdicion — edge (choose takes precedence over edit)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modSuministradorHelper.Suministrador_DobleClick_AbrirEdicion( _
        True, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("action")) <> "choose" Then
        Err.Raise 1002, , "expected action=choose (cmdElegir visible takes precedence), got " & parsed("payload")("action")
    End If

    Test_SuministradorHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence = BuildOk("choose-precedence", logs)
    Exit Function
EH:
    Test_SuministradorHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 22. RunAll — wrapper for Dysflow manifest discovery (access-vba-tdd §1.1.1)
' ---------------------------------------------------------------------------
Public Function Test_SuministradorHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_SuministradorHelper_Abrir_Inicializar_HappyAdminNoArgs", _
        "Test_SuministradorHelper_Abrir_Inicializar_HappyNonAdminNoArgs", _
        "Test_SuministradorHelper_Abrir_Inicializar_HappyAdminWithArgs", _
        "Test_SuministradorHelper_Buscar_Listar_HappyAllRows", _
        "Test_SuministradorHelper_Buscar_Listar_EdgeEmptyCollection", _
        "Test_SuministradorHelper_Buscar_Listar_HappyFilterByNombre", _
        "Test_SuministradorHelper_Buscar_Listar_HappyFilterByCIF", _
        "Test_SuministradorHelper_Buscar_Listar_EdgeLongFilter", _
        "Test_SuministradorHelper_Buscar_Listar_AdversarialSemicolonInCIF", _
        "Test_SuministradorHelper_Seleccionar_Cargar_HappyAdmin", _
        "Test_SuministradorHelper_Seleccionar_Cargar_HappyNonAdmin", _
        "Test_SuministradorHelper_Seleccionar_Cargar_SadEmptySelection", _
        "Test_SuministradorHelper_Seleccionar_Cargar_SadNotFound", _
        "Test_SuministradorHelper_Eliminar_Borrar_SadCancelled", _
        "Test_SuministradorHelper_Eliminar_Borrar_HappyDeleted", _
        "Test_SuministradorHelper_Eliminar_Borrar_SadNoEntity", _
        "Test_SuministradorHelper_Eliminar_Borrar_SadMissingID", _
        "Test_SuministradorHelper_DobleClick_AbrirEdicion_HappyChoose", _
        "Test_SuministradorHelper_DobleClick_AbrirEdicion_HappyEdit", _
        "Test_SuministradorHelper_DobleClick_AbrirEdicion_HappyNone", _
        "Test_SuministradorHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence" _
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
        Test_SuministradorHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_SuministradorHelper_RunAll = BuildOk("all-passed", logs)
    End If
    On Error Resume Next
    Call TeardownFixture
    On Error GoTo 0
    Exit Function
EH:
    On Error Resume Next
    Dim p_Error As String
    p_Error = "Test_SuministradorHelper_RunAll EH: " & Err.Description
    Call TeardownFixture
    On Error GoTo 0
    Test_SuministradorHelper_RunAll = BuildFail(p_Error, logs)
End Function
