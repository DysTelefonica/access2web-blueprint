Attribute VB_Name = "Test_RACSHelper"
Option Compare Database
Option Explicit

' Test_RACSHelper — REWORK TDD atoms for modRACSHelper.bas
' (Form_FormRACSGestion, see docs/audit/racs-gestion-thin.md).
'
' Anti-pattern removed (was in PR #29 commit 979a57a):
'   - Test_RACSHelper_OpenForm(ByRef p_Form, ByRef p_Error) used to call
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

' === Module-level constants (all at top per vba-access §10.1) =======================

' Fixture ID base — entities created by atoms use IDs in this range so teardown is
' surgical (DELETE WHERE IDRAC >= TEST_BASE_ID).
Private Const TEST_BASE_ID As Long = 900800

' Field separator used in rowSource assertions. Mirrors the helper's constant.
Private Const TEST_FIELD_SEP As String = ";"


' === Local helpers (all at top per vba-access §10.1) ================================

' --- BuildOk / BuildFail --------------------------------------------------------------
Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

' --- BuildStubRAC --------------------------------------------------------------------
' Builds a Scripting.Dictionary that mimics the shape of a RAC entity.
' Returns a Dictionary {IDRAC, RAC, CORREO, DESCRIPCION}.
Private Function BuildStubRAC( _
    ByVal p_ID As String, _
    ByVal p_RAC As String, _
    ByVal p_Correo As String, _
    ByVal p_Desc As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("IDRAC") = p_ID
    d("RAC") = p_RAC
    d("CORREO") = p_Correo
    d("DESCRIPCION") = p_Desc
    Set BuildStubRAC = d
End Function

' --- BuildStubRACsDict ---------------------------------------------------------------
' Builds a Dictionary {CStr(ID) -> stubEntity} from an array of names.
' IDs are TEST_BASE_ID, TEST_BASE_ID+1, ... so each row has a unique key.
Private Function BuildStubRACsDict(ByVal p_Names As Variant) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(p_Names) To UBound(p_Names)
        Dim idStr As String
        idStr = CStr(TEST_BASE_ID + i)
        Set d(idStr) = BuildStubRAC(idStr, CStr(p_Names(i)), _
                                    "test-" & i & "@example.com", _
                                    "test-desc-" & i)
    Next i
    Set BuildStubRACsDict = d
End Function

' --- InsertRealRow -------------------------------------------------------------------
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
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_Abrir_Inicializar_HappyAdminNoArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_Abrir_Inicializar(True, False, p_Error)
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

    Test_RACSHelper_Abrir_Inicializar_HappyAdminNoArgs = BuildOk("admin-no-args", logs)
    Exit Function
EH:
    Test_RACSHelper_Abrir_Inicializar_HappyAdminNoArgs = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 2. Abrir_Inicializar — happy (admin=false, no OpenArgs)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_Abrir_Inicializar_HappyNonAdminNoArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_Abrir_Inicializar(False, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("showAlta") <> False Then Err.Raise 1002, , "expected showAlta=false"
    If parsed("payload")("focusAlta") <> False Then Err.Raise 1002, , "expected focusAlta=false"
    If parsed("payload")("showElegir") <> False Then Err.Raise 1002, , "expected showElegir=false"

    Test_RACSHelper_Abrir_Inicializar_HappyNonAdminNoArgs = BuildOk("non-admin-no-args", logs)
    Exit Function
EH:
    Test_RACSHelper_Abrir_Inicializar_HappyNonAdminNoArgs = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 3. Abrir_Inicializar — happy (admin=true, has OpenArgs)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_Abrir_Inicializar_HappyAdminWithArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_Abrir_Inicializar(True, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("showAlta") <> True Then Err.Raise 1002, , "expected showAlta=true"
    If parsed("payload")("focusAlta") <> True Then Err.Raise 1002, , "expected focusAlta=true"
    If parsed("payload")("showElegir") <> True Then Err.Raise 1002, , "expected showElegir=true"

    Test_RACSHelper_Abrir_Inicializar_HappyAdminWithArgs = BuildOk("admin-with-args", logs)
    Exit Function
EH:
    Test_RACSHelper_Abrir_Inicializar_HappyAdminWithArgs = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 4. Buscar_Listar — happy (3 entities, no filter -> 3 rows)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_Buscar_Listar_HappyAllRows() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubRACsDict(Array("ALPHA", "BETA", "GAMMA"))

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_Buscar_Listar(col, "", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    Dim msg As String
    If Not AssertPayloadHasKey(parsed("payload"), "count", msg) Then Err.Raise 1001, , msg
    If Not AssertPayloadHasKey(parsed("payload"), "rowSource", msg) Then Err.Raise 1001, , msg

    If CLng(parsed("payload")("count")) <> 3 Then Err.Raise 1002, , "expected count=3, got " & parsed("payload")("count")

    Dim rs As String
    rs = CStr(parsed("payload")("rowSource"))
    If InStr(rs, "ALPHA") = 0 Then Err.Raise 1003, , "rowSource missing ALPHA"
    If InStr(rs, "BETA") = 0 Then Err.Raise 1003, , "rowSource missing BETA"
    If InStr(rs, "GAMMA") = 0 Then Err.Raise 1003, , "rowSource missing GAMMA"

    Test_RACSHelper_Buscar_Listar_HappyAllRows = BuildOk(3, logs)
    Exit Function
EH:
    Test_RACSHelper_Buscar_Listar_HappyAllRows = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 5. Buscar_Listar — edge (empty collection -> 0 rows)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_Buscar_Listar_EdgeEmptyCollection() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_Buscar_Listar(Nothing, "", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 0 Then Err.Raise 1002, , "expected count=0"

    Test_RACSHelper_Buscar_Listar_EdgeEmptyCollection = BuildOk(0, logs)
    Exit Function
EH:
    Test_RACSHelper_Buscar_Listar_EdgeEmptyCollection = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 6. Buscar_Listar — happy (filter matches subset)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_Buscar_Listar_HappyFilteredRows() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubRACsDict(Array("ALPHA-A", "BETA", "ALPHA-B"))

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_Buscar_Listar(col, "ALPHA", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 2 Then Err.Raise 1002, , "expected count=2 (ALPHA match), got " & parsed("payload")("count")

    Dim rs As String
    rs = CStr(parsed("payload")("rowSource"))
    If InStr(rs, "BETA") > 0 Then Err.Raise 1003, , "rowSource should NOT contain BETA"

    Test_RACSHelper_Buscar_Listar_HappyFilteredRows = BuildOk(2, logs)
    Exit Function
EH:
    Test_RACSHelper_Buscar_Listar_HappyFilteredRows = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 7. Buscar_Listar — edge (256-char filter that does not match anything -> 0 rows)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_Buscar_Listar_EdgeLongFilter() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubRACsDict(Array("SHORT"))

    Dim longFilter As String
    longFilter = String(256, "X")

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_Buscar_Listar(col, longFilter, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 0 Then Err.Raise 1002, , "expected count=0 for 256-char non-matching filter"

    Test_RACSHelper_Buscar_Listar_EdgeLongFilter = BuildOk(0, logs)
    Exit Function
EH:
    Test_RACSHelper_Buscar_Listar_EdgeLongFilter = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 8. Buscar_Listar — adversarial (filter text contains semicolons — must not break)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_Buscar_Listar_AdversarialFilterWithSemicolons() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubRACsDict(Array("SAFE"))

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_Buscar_Listar(col, ";;;;", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 0 Then Err.Raise 1002, , "expected count=0 for semicolon-only filter"

    Test_RACSHelper_Buscar_Listar_AdversarialFilterWithSemicolons = BuildOk(0, logs)
    Exit Function
EH:
    Test_RACSHelper_Buscar_Listar_AdversarialFilterWithSemicolons = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 9. Buscar_Listar — adversarial (semicolons in RAC name are sanitized)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_Buscar_Listar_AdversarialSemicolonInName() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = CreateObject("Scripting.Dictionary")
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("IDRAC") = CStr(TEST_BASE_ID + 1)
    d("RAC") = "A;B;C"
    d("CORREO") = "evil@x.com"
    d("DESCRIPCION") = "evil;desc"
    Set col(CStr(TEST_BASE_ID + 1)) = d

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_Buscar_Listar(col, "", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    Dim rs As String
    rs = CStr(parsed("payload")("rowSource"))
    If InStr(rs, "A;B;C") > 0 Then Err.Raise 1002, , "expected semicolons sanitized"
    If InStr(rs, "A:B:C") = 0 Then Err.Raise 1003, , "expected sanitized value A:B:C"

    Test_RACSHelper_Buscar_Listar_AdversarialSemicolonInName = BuildOk(1, logs)
    Exit Function
EH:
    Test_RACSHelper_Buscar_Listar_AdversarialSemicolonInName = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 10. Seleccionar_Cargar — happy (selection present, admin)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_Seleccionar_Cargar_HappyAdmin() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubRACsDict(Array("FIRST"))

    Dim selectedId As String
    selectedId = CStr(TEST_BASE_ID + 1)

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_Seleccionar_Cargar(selectedId, col, True, p_Error)
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
    If CStr(parsed("payload")("entity")("IDRAC")) <> selectedId Then
        Err.Raise 1005, , "expected ID=" & selectedId & ", got " & parsed("payload")("entity")("IDRAC")
    End If
    If CStr(parsed("payload")("entity")("RAC")) <> "FIRST" Then
        Err.Raise 1006, , "expected name=FIRST, got " & parsed("payload")("entity")("RAC")
    End If

    Test_RACSHelper_Seleccionar_Cargar_HappyAdmin = BuildOk(selectedId, logs)
    Exit Function
EH:
    Test_RACSHelper_Seleccionar_Cargar_HappyAdmin = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 11. Seleccionar_Cargar — happy (selection present, non-admin)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_Seleccionar_Cargar_HappyNonAdmin() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubRACsDict(Array("FIRST"))

    Dim selectedId As String
    selectedId = CStr(TEST_BASE_ID + 1)

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_Seleccionar_Cargar(selectedId, col, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("enableEditar") <> True Then Err.Raise 1002, , "expected enableEditar=true"
    If parsed("payload")("enableEliminar") <> False Then Err.Raise 1003, , "expected enableEliminar=false for non-admin"

    Test_RACSHelper_Seleccionar_Cargar_HappyNonAdmin = BuildOk(selectedId, logs)
    Exit Function
EH:
    Test_RACSHelper_Seleccionar_Cargar_HappyNonAdmin = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 12. Seleccionar_Cargar — sad (empty selection -> nothing selected, both disabled)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_Seleccionar_Cargar_SadEmptySelection() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubRACsDict(Array("FIRST"))

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_Seleccionar_Cargar("", col, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("enableEditar") <> False Then Err.Raise 1002, , "expected enableEditar=false"
    If parsed("payload")("enableEliminar") <> False Then Err.Raise 1003, , "expected enableEliminar=false"
    If Not parsed("payload")("entity") Is Nothing Then Err.Raise 1004, , "expected entity=null"

    Test_RACSHelper_Seleccionar_Cargar_SadEmptySelection = BuildOk("empty", logs)
    Exit Function
EH:
    Test_RACSHelper_Seleccionar_Cargar_SadEmptySelection = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 13. Seleccionar_Cargar — sad (selected ID not found in collection -> entity=null)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_Seleccionar_Cargar_SadNotFound() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubRACsDict(Array("FIRST"))

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_Seleccionar_Cargar("999999", col, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("enableEditar") <> False Then Err.Raise 1002, , "expected enableEditar=false"
    If Not parsed("payload")("entity") Is Nothing Then Err.Raise 1004, , "expected entity=null"

    Test_RACSHelper_Seleccionar_Cargar_SadNotFound = BuildOk("not-found", logs)
    Exit Function
EH:
    Test_RACSHelper_Seleccionar_Cargar_SadNotFound = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 14. Eliminar_Borrar — sad (p_PromptResult=vbNo -> cancelled, row preserved)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_Eliminar_Borrar_SadCancelled() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    On Error GoTo EH

    Dim seedId As String
    seedId = CStr(TEST_BASE_ID + 1)
    Dim insertErr As String
    If Not InsertRealRow(seedId, "CANCEL-TEST", "cancel@example.com", "desc", insertErr) Then
        Err.Raise 1001, , insertErr
    End If

    Dim entity As Object
    Set entity = BuildStubRAC(seedId, "CANCEL-TEST", "cancel@example.com", "desc")

    Dim promptResult As Long
    promptResult = vbNo

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_Eliminar_Borrar(entity, promptResult, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If InStr(json, """ok"":true") = 0 Then Err.Raise 1002, , "expected ok=true for cancelled path, got " & json

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim countAfter As Long
    countAfter = DCount("*", "TbRACS", "IDRAC=" & seedId)
    If countAfter <> 1 Then Err.Raise 1003, , "row should still exist after cancel, count=" & countAfter

    Test_RACSHelper_Eliminar_Borrar_SadCancelled = BuildOk("cancelled", logs)
    Call TeardownFixture
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    On Error GoTo 0
    Test_RACSHelper_Eliminar_Borrar_SadCancelled = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 15. Eliminar_Borrar — happy (p_PromptResult=vbYes -> deleted, row gone)
'     Cardinalidad before/after (access-vba-tdd §4.5).
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_Eliminar_Borrar_HappyDeleted() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    On Error GoTo EH

    Dim seedId As String
    seedId = CStr(TEST_BASE_ID + 1)
    Dim insertErr As String
    If Not InsertRealRow(seedId, "DELETE-TEST", "del@example.com", "desc", insertErr) Then
        Err.Raise 1001, , insertErr
    End If

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim countBefore As Long
    countBefore = DCount("*", "TbRACS", "IDRAC=" & seedId)
    If countBefore <> 1 Then Err.Raise 1002, , "expected 1 row before delete, got " & countBefore

    Dim entity As Object
    Set entity = BuildStubRAC(seedId, "DELETE-TEST", "del@example.com", "desc")

    Dim promptResult As Long
    promptResult = vbYes

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_Eliminar_Borrar(entity, promptResult, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If InStr(json, """ok"":true") = 0 Then Err.Raise 1003, , "expected ok=true for deleted path, got " & json

    Dim countAfter As Long
    countAfter = DCount("*", "TbRACS", "IDRAC=" & seedId)
    If countAfter <> 0 Then Err.Raise 1004, , "expected 0 rows after delete, got " & countAfter

    logs(0) = "Eliminar_Borrar_HappyDeleted: cardinality " & countBefore & "->" & countAfter

    Test_RACSHelper_Eliminar_Borrar_HappyDeleted = BuildOk("deleted", logs)
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    On Error GoTo 0
    Test_RACSHelper_Eliminar_Borrar_HappyDeleted = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 16. Eliminar_Borrar — sad (no entity supplied -> fail JSON)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_Eliminar_Borrar_SadNoEntity() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_Eliminar_Borrar(Nothing, vbYes, p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1002, , "expected ok=false, got " & json

    Test_RACSHelper_Eliminar_Borrar_SadNoEntity = BuildOk("no-entity-rejected", logs)
    Exit Function
EH:
    Test_RACSHelper_Eliminar_Borrar_SadNoEntity = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 17. DobleClick_AbrirEdicion — happy (choose dispatched when cmdElegir visible)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_DobleClick_AbrirEdicion_HappyChoose() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_DobleClick_AbrirEdicion(True, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("action")) <> "choose" Then
        Err.Raise 1002, , "expected action=choose, got " & parsed("payload")("action")
    End If

    Test_RACSHelper_DobleClick_AbrirEdicion_HappyChoose = BuildOk("choose", logs)
    Exit Function
EH:
    Test_RACSHelper_DobleClick_AbrirEdicion_HappyChoose = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 18. DobleClick_AbrirEdicion — happy (edit dispatched when ComandoEditar enabled)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_DobleClick_AbrirEdicion_HappyEdit() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_DobleClick_AbrirEdicion(False, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("action")) <> "edit" Then
        Err.Raise 1002, , "expected action=edit, got " & parsed("payload")("action")
    End If

    Test_RACSHelper_DobleClick_AbrirEdicion_HappyEdit = BuildOk("edit", logs)
    Exit Function
EH:
    Test_RACSHelper_DobleClick_AbrirEdicion_HappyEdit = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 19. DobleClick_AbrirEdicion — happy (none when both buttons unavailable)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_DobleClick_AbrirEdicion_HappyNone() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_DobleClick_AbrirEdicion(False, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("action")) <> "none" Then
        Err.Raise 1002, , "expected action=none, got " & parsed("payload")("action")
    End If

    Test_RACSHelper_DobleClick_AbrirEdicion_HappyNone = BuildOk("none", logs)
    Exit Function
EH:
    Test_RACSHelper_DobleClick_AbrirEdicion_HappyNone = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 20. DobleClick_AbrirEdicion — edge (choose takes precedence over edit)
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modRACSHelper.RACS_DobleClick_AbrirEdicion(True, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("action")) <> "choose" Then
        Err.Raise 1002, , "expected action=choose (cmdElegir visible takes precedence), got " & parsed("payload")("action")
    End If

    Test_RACSHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence = BuildOk("choose-precedence", logs)
    Exit Function
EH:
    Test_RACSHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 21. RunAll — wrapper for Dysflow manifest discovery
' ---------------------------------------------------------------------------
Public Function Test_RACSHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_RACSHelper_Abrir_Inicializar_HappyAdminNoArgs", _
        "Test_RACSHelper_Abrir_Inicializar_HappyNonAdminNoArgs", _
        "Test_RACSHelper_Abrir_Inicializar_HappyAdminWithArgs", _
        "Test_RACSHelper_Buscar_Listar_HappyAllRows", _
        "Test_RACSHelper_Buscar_Listar_EdgeEmptyCollection", _
        "Test_RACSHelper_Buscar_Listar_HappyFilteredRows", _
        "Test_RACSHelper_Buscar_Listar_EdgeLongFilter", _
        "Test_RACSHelper_Buscar_Listar_AdversarialFilterWithSemicolons", _
        "Test_RACSHelper_Buscar_Listar_AdversarialSemicolonInName", _
        "Test_RACSHelper_Seleccionar_Cargar_HappyAdmin", _
        "Test_RACSHelper_Seleccionar_Cargar_HappyNonAdmin", _
        "Test_RACSHelper_Seleccionar_Cargar_SadEmptySelection", _
        "Test_RACSHelper_Seleccionar_Cargar_SadNotFound", _
        "Test_RACSHelper_Eliminar_Borrar_SadCancelled", _
        "Test_RACSHelper_Eliminar_Borrar_HappyDeleted", _
        "Test_RACSHelper_Eliminar_Borrar_SadNoEntity", _
        "Test_RACSHelper_DobleClick_AbrirEdicion_HappyChoose", _
        "Test_RACSHelper_DobleClick_AbrirEdicion_HappyEdit", _
        "Test_RACSHelper_DobleClick_AbrirEdicion_HappyNone", _
        "Test_RACSHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence" _
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
        Test_RACSHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_RACSHelper_RunAll = BuildOk("all-passed", logs)
    End If
    On Error Resume Next
    Call TeardownFixture
    On Error GoTo 0
    Exit Function
EH:
    On Error Resume Next
    Dim p_Error As String
    p_Error = "Test_RACSHelper_RunAll EH: " & Err.Description
    Call TeardownFixture
    On Error GoTo 0
    Test_RACSHelper_RunAll = BuildFail(p_Error, logs)
End Function
