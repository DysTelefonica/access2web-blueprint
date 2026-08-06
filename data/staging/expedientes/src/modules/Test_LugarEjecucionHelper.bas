Attribute VB_Name = "Test_LugarEjecucionHelper"
Option Compare Database
Option Explicit

' Test_LugarEjecucionHelper — REWORK TDD atoms for modLugarEjecucionHelper.bas
' (Form_FormLugarEjecucionGestion, see docs/audit/lugar-ejecucion-gestion-thin.md).
'
' Anti-pattern removed (was in PR #29 commit 979a57a):
'   - Test_LugarEjecucionHelper_OpenForm(ByRef p_Form, ByRef p_Error) used to call
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
' surgical (DELETE WHERE IDLugarEjecucion >= TEST_BASE_ID).
Private Const TEST_BASE_ID As Long = 900600

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

' --- BuildStubLugarEjecucion -----------------------------------------------------------
' Builds a Scripting.Dictionary that mimics the shape of a LugarEjecucion entity.
' The helper accepts p_Lugares as a Dictionary, so the stub is a pure data structure.
' Returns a Dictionary {IDLugarEjecucion, LugarEjecucion, DESCRIPCION}.
Private Function BuildStubLugarEjecucion( _
    ByVal p_ID As String, _
    ByVal p_Lugar As String, _
    ByVal p_Desc As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("IDLugarEjecucion") = p_ID
    d("LugarEjecucion") = p_Lugar
    d("DESCRIPCION") = p_Desc
    Set BuildStubLugarEjecucion = d
End Function

' --- BuildStubLugaresDict ------------------------------------------------------------
' Builds a Dictionary {CStr(ID) -> stubEntity} from an array of names.
' IDs are TEST_BASE_ID, TEST_BASE_ID+1, ... so each row has a unique key.
Private Function BuildStubLugaresDict(ByVal p_Names As Variant) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(p_Names) To UBound(p_Names)
        Dim idStr As String
        idStr = CStr(TEST_BASE_ID + i)
        Set d(idStr) = BuildStubLugarEjecucion(idStr, CStr(p_Names(i)), "Desc " & CStr(p_Names(i)))
    Next i
    Set BuildStubLugaresDict = d
End Function

' --- InsertRealRow -------------------------------------------------------------------
' Inserts a row into TbLugaresEjecucion for the Eliminar_Borrar happy/adversarial tests
' (which exercise the DAO path via Helper_EntidadCRUD). Returns True on success.
Private Function InsertRealRow( _
    ByVal p_ID As String, _
    ByVal p_Lugar As String, _
    ByVal p_Desc As String, _
    ByRef p_Error As String _
) As Boolean
    On Error GoTo EH
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim sql As String
    sql = "INSERT INTO TbLugaresEjecucion (IDLugarEjecucion, LugarEjecucion, Descripcion) VALUES (" & _
          p_ID & ", '" & Replace(p_Lugar, "'", "''") & "', '" & Replace(p_Desc, "'", "''") & "')"
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
    before = DCount("*", "TbLugaresEjecucion", "IDLugarEjecucion >= " & TEST_BASE_ID)
    db.Execute "DELETE FROM TbLugaresEjecucion WHERE IDLugarEjecucion >= " & TEST_BASE_ID, dbFailOnError
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
' 1. Abrir_Inicializar — happy (admin=true, no OpenArgs)
'    Expected: showAlta=true, focusAlta=true, showElegir=false
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_Abrir_Inicializar_HappyAdminNoArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Abrir_Inicializar( _
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

    Test_LugarEjecucionHelper_Abrir_Inicializar_HappyAdminNoArgs = BuildOk("admin-no-args", logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_Abrir_Inicializar_HappyAdminNoArgs = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 2. Abrir_Inicializar — happy (admin=false, no OpenArgs)
'    Expected: showAlta=false, focusAlta=false, showElegir=false
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_Abrir_Inicializar_HappyNonAdminNoArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Abrir_Inicializar( _
        False, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("showAlta") <> False Then Err.Raise 1002, , "expected showAlta=false"
    If parsed("payload")("focusAlta") <> False Then Err.Raise 1002, , "expected focusAlta=false"
    If parsed("payload")("showElegir") <> False Then Err.Raise 1002, , "expected showElegir=false"

    Test_LugarEjecucionHelper_Abrir_Inicializar_HappyNonAdminNoArgs = BuildOk("non-admin-no-args", logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_Abrir_Inicializar_HappyNonAdminNoArgs = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 3. Abrir_Inicializar — happy (admin=true, has OpenArgs)
'    Expected: showAlta=true, focusAlta=true, showElegir=true
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_Abrir_Inicializar_HappyAdminWithArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Abrir_Inicializar( _
        True, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("showAlta") <> True Then Err.Raise 1002, , "expected showAlta=true"
    If parsed("payload")("focusAlta") <> True Then Err.Raise 1002, , "expected focusAlta=true"
    If parsed("payload")("showElegir") <> True Then Err.Raise 1002, , "expected showElegir=true"

    Test_LugarEjecucionHelper_Abrir_Inicializar_HappyAdminWithArgs = BuildOk("admin-with-args", logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_Abrir_Inicializar_HappyAdminWithArgs = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 4. Buscar_Listar — happy (3 entities, no filter -> 3 rows)
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_Buscar_Listar_HappyAllRows() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubLugaresDict(Array("OFICINA_A", "OFICINA_B", "OFICINA_C"))

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Buscar_Listar(col, "", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim msg As String
    If Not AssertPayloadHasKey(parsed("payload"), "count", msg) Then Err.Raise 1001, , msg
    If Not AssertPayloadHasKey(parsed("payload"), "rowSource", msg) Then Err.Raise 1001, , msg

    If CLng(parsed("payload")("count")) <> 3 Then Err.Raise 1002, , "expected count=3, got " & parsed("payload")("count")

    Dim rs As String
    rs = CStr(parsed("payload")("rowSource"))
    If InStr(rs, "OFICINA_A") = 0 Then Err.Raise 1003, , "rowSource missing OFICINA_A"
    If InStr(rs, "OFICINA_B") = 0 Then Err.Raise 1003, , "rowSource missing OFICINA_B"
    If InStr(rs, "OFICINA_C") = 0 Then Err.Raise 1003, , "rowSource missing OFICINA_C"

    Test_LugarEjecucionHelper_Buscar_Listar_HappyAllRows = BuildOk(3, logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_Buscar_Listar_HappyAllRows = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 5. Buscar_Listar — edge (empty collection -> 0 rows)
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_Buscar_Listar_EdgeEmptyCollection() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Buscar_Listar(Nothing, "", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 0 Then Err.Raise 1002, , "expected count=0"

    Test_LugarEjecucionHelper_Buscar_Listar_EdgeEmptyCollection = BuildOk(0, logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_Buscar_Listar_EdgeEmptyCollection = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 6. Buscar_Listar — happy (filter matches subset)
'    Names: OFICINA-A, OFICINA_B, OFICINA-C; filter "OFICINA" -> 3 rows
'    Names: SEDE_A, OFICINA_B, SEDE_C; filter "OFICINA" -> 1 row
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_Buscar_Listar_HappyFilteredRows() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubLugaresDict(Array("OFICINA_A", "OFICINA_B", "SEDE_C"))

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Buscar_Listar(col, "OFICINA", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 2 Then Err.Raise 1002, , "expected count=2 (OFICINA match), got " & parsed("payload")("count")

    Dim rs As String
    rs = CStr(parsed("payload")("rowSource"))
    If InStr(rs, "SEDE_C") > 0 Then Err.Raise 1003, , "rowSource should NOT contain SEDE_C"

    Test_LugarEjecucionHelper_Buscar_Listar_HappyFilteredRows = BuildOk(2, logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_Buscar_Listar_HappyFilteredRows = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 7. Buscar_Listar — edge (256-char filter that does not match anything -> 0 rows)
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_Buscar_Listar_EdgeLongFilter() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubLugaresDict(Array("SHORT"))

    Dim longFilter As String
    longFilter = String(256, "X")

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Buscar_Listar(col, longFilter, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 0 Then Err.Raise 1002, , "expected count=0 for 256-char non-matching filter"

    Test_LugarEjecucionHelper_Buscar_Listar_EdgeLongFilter = BuildOk(0, logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_Buscar_Listar_EdgeLongFilter = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 8. Buscar_Listar — adversarial (filter text contains semicolons — must not break)
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_Buscar_Listar_AdversarialFilterWithSemicolons() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubLugaresDict(Array("SAFE"))

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Buscar_Listar(col, ";;;;", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 0 Then Err.Raise 1002, , "expected count=0 for semicolon-only filter"

    Test_LugarEjecucionHelper_Buscar_Listar_AdversarialFilterWithSemicolons = BuildOk(0, logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_Buscar_Listar_AdversarialFilterWithSemicolons = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 9. Buscar_Listar — adversarial (semicolons in LugarEjecucion name field are sanitized to colons)
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_Buscar_Listar_AdversarialSemicolonInName() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    ' Manually build a Dictionary with a name containing semicolons.
    Dim col As Object
    Set col = CreateObject("Scripting.Dictionary")
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("IDLugarEjecucion") = CStr(TEST_BASE_ID + 1)
    d("LugarEjecucion") = "A;B;C"
    d("DESCRIPCION") = "evil;desc"
    Set col(CStr(TEST_BASE_ID + 1)) = d

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Buscar_Listar(col, "", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    Dim rs As String
    rs = CStr(parsed("payload")("rowSource"))
    ' The helper must NOT include literal semicolons in the field values — they get
    ' replaced with colons to keep the rowSource well-formed.
    If InStr(rs, "A;B;C") > 0 Then Err.Raise 1002, , "expected semicolons sanitized"
    If InStr(rs, "A:B:C") = 0 Then Err.Raise 1003, , "expected sanitized value A:B:C"

    Test_LugarEjecucionHelper_Buscar_Listar_AdversarialSemicolonInName = BuildOk(1, logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_Buscar_Listar_AdversarialSemicolonInName = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 10. Seleccionar_Cargar — happy (selection present, admin)
'     Expected: entity=stub, enableEditar=true, enableEliminar=true
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_Seleccionar_Cargar_HappyAdmin() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubLugaresDict(Array("FIRST"))

    Dim selectedId As String
    selectedId = CStr(TEST_BASE_ID + 1)

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Seleccionar_Cargar( _
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
    If CStr(parsed("payload")("entity")("IDLugarEjecucion")) <> selectedId Then
        Err.Raise 1005, , "expected ID=" & selectedId & ", got " & parsed("payload")("entity")("IDLugarEjecucion")
    End If
    If CStr(parsed("payload")("entity")("LugarEjecucion")) <> "FIRST" Then
        Err.Raise 1006, , "expected name=FIRST, got " & parsed("payload")("entity")("LugarEjecucion")
    End If

    Test_LugarEjecucionHelper_Seleccionar_Cargar_HappyAdmin = BuildOk(selectedId, logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_Seleccionar_Cargar_HappyAdmin = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 11. Seleccionar_Cargar — happy (selection present, non-admin)
'     Expected: enableEditar=true, enableEliminar=false
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_Seleccionar_Cargar_HappyNonAdmin() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubLugaresDict(Array("FIRST"))

    Dim selectedId As String
    selectedId = CStr(TEST_BASE_ID + 1)

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Seleccionar_Cargar( _
        selectedId, col, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("enableEditar") <> True Then Err.Raise 1002, , "expected enableEditar=true"
    If parsed("payload")("enableEliminar") <> False Then Err.Raise 1003, , "expected enableEliminar=false for non-admin"

    Test_LugarEjecucionHelper_Seleccionar_Cargar_HappyNonAdmin = BuildOk(selectedId, logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_Seleccionar_Cargar_HappyNonAdmin = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 12. Seleccionar_Cargar — sad (empty selection -> nothing selected, both disabled)
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_Seleccionar_Cargar_SadEmptySelection() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubLugaresDict(Array("FIRST"))

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Seleccionar_Cargar( _
        "", col, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("enableEditar") <> False Then Err.Raise 1002, , "expected enableEditar=false"
    If parsed("payload")("enableEliminar") <> False Then Err.Raise 1003, , "expected enableEliminar=false"
    If Not parsed("payload")("entity") Is Nothing Then Err.Raise 1004, , "expected entity=null"

    Test_LugarEjecucionHelper_Seleccionar_Cargar_SadEmptySelection = BuildOk("empty", logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_Seleccionar_Cargar_SadEmptySelection = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 13. Seleccionar_Cargar — sad (selected ID not found in collection -> entity=null, both disabled)
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_Seleccionar_Cargar_SadNotFound() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubLugaresDict(Array("FIRST"))

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Seleccionar_Cargar( _
        "999999", col, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("enableEditar") <> False Then Err.Raise 1002, , "expected enableEditar=false"
    If Not parsed("payload")("entity") Is Nothing Then Err.Raise 1004, , "expected entity=null"

    Test_LugarEjecucionHelper_Seleccionar_Cargar_SadNotFound = BuildOk("not-found", logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_Seleccionar_Cargar_SadNotFound = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 14. Eliminar_Borrar — sad (p_PromptResult=vbNo -> cancelled, row preserved)
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_Eliminar_Borrar_SadCancelled() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    On Error GoTo EH

    Dim seedId As String
    seedId = CStr(TEST_BASE_ID + 1)
    Dim insertErr As String
    If Not InsertRealRow(seedId, "CANCEL-TEST", "desc", insertErr) Then
        Err.Raise 1001, , insertErr
    End If

    Dim entity As Object
    Set entity = BuildStubLugarEjecucion(seedId, "CANCEL-TEST", "desc")

    Dim promptResult As Long
    promptResult = vbNo

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Eliminar_Borrar(entity, promptResult, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    ' Helper returns ok=true with no payload for "cancelled" path
    If InStr(json, """ok"":true") = 0 Then Err.Raise 1002, , "expected ok=true for cancelled path, got " & json

    ' Cardinalidad: row must still exist.
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim countAfter As Long
    countAfter = DCount("*", "TbLugaresEjecucion", "IDLugarEjecucion=" & seedId)
    If countAfter <> 1 Then Err.Raise 1003, , "row should still exist after cancel, count=" & countAfter

    Test_LugarEjecucionHelper_Eliminar_Borrar_SadCancelled = BuildOk("cancelled", logs)
    Call TeardownFixture
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    On Error GoTo 0
    Test_LugarEjecucionHelper_Eliminar_Borrar_SadCancelled = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 15. Eliminar_Borrar — happy (p_PromptResult=vbYes -> deleted, row gone)
'     Cardinalidad before/after (access-vba-tdd §4.5).
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_Eliminar_Borrar_HappyDeleted() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    Dim seedId As String
    seedId = CStr(TEST_BASE_ID + 1)
    Dim insertErr As String
    If Not InsertRealRow(seedId, "DELETE-TEST", "desc", insertErr) Then
        Err.Raise 1001, , insertErr
    End If

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim countBefore As Long
    countBefore = DCount("*", "TbLugaresEjecucion", "IDLugarEjecucion=" & seedId)
    If countBefore <> 1 Then Err.Raise 1002, , "expected 1 row before delete, got " & countBefore

    Dim entity As Object
    Set entity = BuildStubLugarEjecucion(seedId, "DELETE-TEST", "desc")

    Dim promptResult As Long
    promptResult = vbYes

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Eliminar_Borrar(entity, promptResult, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If InStr(json, """ok"":true") = 0 Then Err.Raise 1003, , "expected ok=true for deleted path, got " & json

    Dim countAfter As Long
    countAfter = DCount("*", "TbLugaresEjecucion", "IDLugarEjecucion=" & seedId)
    If countAfter <> 0 Then Err.Raise 1004, , "expected 0 rows after delete, got " & countAfter

    logs(0) = "Eliminar_Borrar_HappyDeleted: cardinality " & countBefore & "->" & countAfter

    Test_LugarEjecucionHelper_Eliminar_Borrar_HappyDeleted = BuildOk("deleted", logs)
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    On Error GoTo 0
    Test_LugarEjecucionHelper_Eliminar_Borrar_HappyDeleted = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 16. Eliminar_Borrar — sad (no entity supplied -> fail JSON)
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_Eliminar_Borrar_SadNoEntity() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_Eliminar_Borrar(Nothing, vbYes, p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false, got " & json

    Test_LugarEjecucionHelper_Eliminar_Borrar_SadNoEntity = BuildOk("no-entity-rejected", logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_Eliminar_Borrar_SadNoEntity = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 17. DobleClick_AbrirEdicion — happy (choose dispatched when cmdElegir visible)
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_DobleClick_AbrirEdicion_HappyChoose() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_DobleClick_AbrirEdicion( _
        True, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("action")) <> "choose" Then
        Err.Raise 1002, , "expected action=choose, got " & parsed("payload")("action")
    End If

    Test_LugarEjecucionHelper_DobleClick_AbrirEdicion_HappyChoose = BuildOk("choose", logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_DobleClick_AbrirEdicion_HappyChoose = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 18. DobleClick_AbrirEdicion — happy (edit dispatched when ComandoEditar enabled)
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_DobleClick_AbrirEdicion_HappyEdit() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_DobleClick_AbrirEdicion( _
        False, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("action")) <> "edit" Then
        Err.Raise 1002, , "expected action=edit, got " & parsed("payload")("action")
    End If

    Test_LugarEjecucionHelper_DobleClick_AbrirEdicion_HappyEdit = BuildOk("edit", logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_DobleClick_AbrirEdicion_HappyEdit = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 19. DobleClick_AbrirEdicion — happy (none when both buttons unavailable)
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_DobleClick_AbrirEdicion_HappyNone() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_DobleClick_AbrirEdicion( _
        False, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("action")) <> "none" Then
        Err.Raise 1002, , "expected action=none, got " & parsed("payload")("action")
    End If

    Test_LugarEjecucionHelper_DobleClick_AbrirEdicion_HappyNone = BuildOk("none", logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_DobleClick_AbrirEdicion_HappyNone = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 20. DobleClick_AbrirEdicion — edge (choose takes precedence over edit)
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    ' Even when edit is enabled, choose (cmdElegir visible) wins.
    Dim p_Error As String
    Dim json As String
    json = modLugarEjecucionHelper.LugarEjecucion_DobleClick_AbrirEdicion( _
        True, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("action")) <> "choose" Then
        Err.Raise 1002, , "expected action=choose (cmdElegir visible takes precedence), got " & parsed("payload")("action")
    End If

    Test_LugarEjecucionHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence = BuildOk("choose-precedence", logs)
    Exit Function
EH:
    Test_LugarEjecucionHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 21. RunAll — wrapper for Dysflow manifest discovery (access-vba-tdd §1.1.1)
' ---------------------------------------------------------------------------
Public Function Test_LugarEjecucionHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_LugarEjecucionHelper_Abrir_Inicializar_HappyAdminNoArgs", _
        "Test_LugarEjecucionHelper_Abrir_Inicializar_HappyNonAdminNoArgs", _
        "Test_LugarEjecucionHelper_Abrir_Inicializar_HappyAdminWithArgs", _
        "Test_LugarEjecucionHelper_Buscar_Listar_HappyAllRows", _
        "Test_LugarEjecucionHelper_Buscar_Listar_EdgeEmptyCollection", _
        "Test_LugarEjecucionHelper_Buscar_Listar_HappyFilteredRows", _
        "Test_LugarEjecucionHelper_Buscar_Listar_EdgeLongFilter", _
        "Test_LugarEjecucionHelper_Buscar_Listar_AdversarialFilterWithSemicolons", _
        "Test_LugarEjecucionHelper_Buscar_Listar_AdversarialSemicolonInName", _
        "Test_LugarEjecucionHelper_Seleccionar_Cargar_HappyAdmin", _
        "Test_LugarEjecucionHelper_Seleccionar_Cargar_HappyNonAdmin", _
        "Test_LugarEjecucionHelper_Seleccionar_Cargar_SadEmptySelection", _
        "Test_LugarEjecucionHelper_Seleccionar_Cargar_SadNotFound", _
        "Test_LugarEjecucionHelper_Eliminar_Borrar_SadCancelled", _
        "Test_LugarEjecucionHelper_Eliminar_Borrar_HappyDeleted", _
        "Test_LugarEjecucionHelper_Eliminar_Borrar_SadNoEntity", _
        "Test_LugarEjecucionHelper_DobleClick_AbrirEdicion_HappyChoose", _
        "Test_LugarEjecucionHelper_DobleClick_AbrirEdicion_HappyEdit", _
        "Test_LugarEjecucionHelper_DobleClick_AbrirEdicion_HappyNone", _
        "Test_LugarEjecucionHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence" _
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
        Test_LugarEjecucionHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_LugarEjecucionHelper_RunAll = BuildOk("all-passed", logs)
    End If
    On Error Resume Next
    Call TeardownFixture
    On Error GoTo 0
    Exit Function
EH:
    On Error Resume Next
    Dim p_Error As String
    p_Error = "Test_LugarEjecucionHelper_RunAll EH: " & Err.Description
    Call TeardownFixture
    On Error GoTo 0
    Test_LugarEjecucionHelper_RunAll = BuildFail(p_Error, logs)
End Function
