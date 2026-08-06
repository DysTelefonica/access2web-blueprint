Attribute VB_Name = "Test_ComercialHelper"
Option Compare Database
Option Explicit

' Test_ComercialHelper — REWORK TDD atoms for modComercialHelper.bas
' (Form_FormComercialesGestion, see docs/audit/comercial-gestion-pure-data.md).
'
' Anti-pattern removed:
'   - Test_ComercialHelper_OpenForm(ByRef p_Form, ByRef p_Error) used to call
'     DoCmd.OpenForm TEST_FORM_NAME and return Forms(TEST_FORM_NAME) as p_Form.
'     This caused VBE UI interruption in headless runs.
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
'   - §1.10 atom signature MUST match helper signature EXACTLY.
'   - §4.2 no-humo: atoms assert concrete values, not "did not crash".
'   - §4.4 strong assertions: count=0 vs >0, payload keys present, etc.
'   - §4.5 cardinalidad for mutaciones: countBefore / countAfter for Eliminar tests.
'   - §5.1 fixture IDs in test range (>= 900000).

' === Module-level constants (all at top per vba-access §10.1) =======================

Private Const TEST_BASE_ID As Long = 900500

Private Const TEST_FIELD_SEP As String = ";"


' === Local helpers (all at top per vba-access §10.1) ================================

Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

' --- BuildStubComercial ----------------------------------------------------------
Private Function BuildStubComercial( _
    ByVal p_ID As String, _
    ByVal p_Nombre As String, _
    ByVal p_Desc As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("IDComercial") = p_ID
    d("Comercial") = p_Nombre
    d("DESCRIPCION") = p_Desc
    Set BuildStubComercial = d
End Function

' --- BuildStubComercialesDict ----------------------------------------------------
Private Function BuildStubComercialesDict(ByVal p_Names As Variant) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(p_Names) To UBound(p_Names)
        Dim idStr As String
        idStr = CStr(TEST_BASE_ID + i)
        Set d(idStr) = BuildStubComercial(idStr, CStr(p_Names(i)), "Desc " & CStr(p_Names(i)))
    Next i
    Set BuildStubComercialesDict = d
End Function

' --- InsertRealRow -------------------------------------------------------------------
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

' --- TeardownFixture ------------------------------------------------------------------
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

' 1. Abrir_Inicializar — happy admin no args
Public Function Test_ComercialHelper_Abrir_Inicializar_HappyAdminNoArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_Abrir_Inicializar(True, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("showAlta") <> True Then Err.Raise 1001, , "expected showAlta=true"
    If parsed("payload")("focusAlta") <> True Then Err.Raise 1002, , "expected focusAlta=true"
    If parsed("payload")("showElegir") <> False Then Err.Raise 1003, , "expected showElegir=false"

    Test_ComercialHelper_Abrir_Inicializar_HappyAdminNoArgs = BuildOk("admin-no-args", logs)
    Exit Function
EH:
    Test_ComercialHelper_Abrir_Inicializar_HappyAdminNoArgs = BuildFail(p_Error, logs)
End Function

' 2. Abrir_Inicializar — happy non-admin no args
Public Function Test_ComercialHelper_Abrir_Inicializar_HappyNonAdminNoArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_Abrir_Inicializar(False, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("showAlta") <> False Then Err.Raise 1001, , "expected showAlta=false"
    If parsed("payload")("focusAlta") <> False Then Err.Raise 1002, , "expected focusAlta=false"
    If parsed("payload")("showElegir") <> False Then Err.Raise 1003, , "expected showElegir=false"

    Test_ComercialHelper_Abrir_Inicializar_HappyNonAdminNoArgs = BuildOk("non-admin-no-args", logs)
    Exit Function
EH:
    Test_ComercialHelper_Abrir_Inicializar_HappyNonAdminNoArgs = BuildFail(p_Error, logs)
End Function

' 3. Abrir_Inicializar — happy admin with args
Public Function Test_ComercialHelper_Abrir_Inicializar_HappyAdminWithArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_Abrir_Inicializar(True, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("showElegir") <> True Then Err.Raise 1001, , "expected showElegir=true"

    Test_ComercialHelper_Abrir_Inicializar_HappyAdminWithArgs = BuildOk("admin-with-args", logs)
    Exit Function
EH:
    Test_ComercialHelper_Abrir_Inicializar_HappyAdminWithArgs = BuildFail(p_Error, logs)
End Function

' 4. Buscar_Listar — happy all rows
Public Function Test_ComercialHelper_Buscar_Listar_HappyAllRows() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubComercialesDict(Array("ALPHA", "BETA", "GAMMA"))

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_Buscar_Listar(col, "", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 3 Then Err.Raise 1002, , "expected count=3, got " & parsed("payload")("count")
    Dim rs As String
    rs = CStr(parsed("payload")("rowSource"))
    If InStr(rs, "ALPHA") = 0 Then Err.Raise 1003, , "rowSource missing ALPHA"
    If InStr(rs, "BETA") = 0 Then Err.Raise 1004, , "rowSource missing BETA"
    If InStr(rs, "GAMMA") = 0 Then Err.Raise 1005, , "rowSource missing GAMMA"

    Test_ComercialHelper_Buscar_Listar_HappyAllRows = BuildOk(3, logs)
    Exit Function
EH:
    Test_ComercialHelper_Buscar_Listar_HappyAllRows = BuildFail(p_Error, logs)
End Function

' 5. Buscar_Listar — edge empty collection
Public Function Test_ComercialHelper_Buscar_Listar_EdgeEmptyCollection() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_Buscar_Listar(Nothing, "", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CLng(parsed("payload")("count")) <> 0 Then Err.Raise 1001, , "expected count=0"

    Test_ComercialHelper_Buscar_Listar_EdgeEmptyCollection = BuildOk(0, logs)
    Exit Function
EH:
    Test_ComercialHelper_Buscar_Listar_EdgeEmptyCollection = BuildFail(p_Error, logs)
End Function

' 6. Buscar_Listar — happy filtered
Public Function Test_ComercialHelper_Buscar_Listar_HappyFilteredRows() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubComercialesDict(Array("ALPHA-A", "BETA", "ALPHA-B"))

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_Buscar_Listar(col, "ALPHA", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CLng(parsed("payload")("count")) <> 2 Then Err.Raise 1001, , "expected count=2 (ALPHA match)"

    Test_ComercialHelper_Buscar_Listar_HappyFilteredRows = BuildOk(2, logs)
    Exit Function
EH:
    Test_ComercialHelper_Buscar_Listar_HappyFilteredRows = BuildFail(p_Error, logs)
End Function

' 7. Buscar_Listar — adversarial filter with semicolons
Public Function Test_ComercialHelper_Buscar_Listar_AdversarialFilterWithSemicolons() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubComercialesDict(Array("SAFE"))

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_Buscar_Listar(col, ";;;;", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CLng(parsed("payload")("count")) <> 0 Then Err.Raise 1001, , "expected count=0 for semicolon-only filter"

    Test_ComercialHelper_Buscar_Listar_AdversarialFilterWithSemicolons = BuildOk(0, logs)
    Exit Function
EH:
    Test_ComercialHelper_Buscar_Listar_AdversarialFilterWithSemicolons = BuildFail(p_Error, logs)
End Function

' 8. Seleccionar_Cargar — happy admin
Public Function Test_ComercialHelper_Seleccionar_Cargar_HappyAdmin() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubComercialesDict(Array("FIRST"))

    Dim selectedId As String
    selectedId = CStr(TEST_BASE_ID + 1)

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_Seleccionar_Cargar(selectedId, col, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("enableEditar") <> True Then Err.Raise 1001, , "expected enableEditar=true"
    If parsed("payload")("enableEliminar") <> True Then Err.Raise 1002, , "expected enableEliminar=true for admin"
    If parsed("payload")("entity") Is Nothing Then Err.Raise 1003, , "expected entity populated"

    Test_ComercialHelper_Seleccionar_Cargar_HappyAdmin = BuildOk(selectedId, logs)
    Exit Function
EH:
    Test_ComercialHelper_Seleccionar_Cargar_HappyAdmin = BuildFail(p_Error, logs)
End Function

' 9. Seleccionar_Cargar — happy non-admin (enableEliminar=false)
Public Function Test_ComercialHelper_Seleccionar_Cargar_HappyNonAdmin() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubComercialesDict(Array("FIRST"))

    Dim selectedId As String
    selectedId = CStr(TEST_BASE_ID + 1)

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_Seleccionar_Cargar(selectedId, col, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("enableEliminar") <> False Then Err.Raise 1001, , "expected enableEliminar=false for non-admin"

    Test_ComercialHelper_Seleccionar_Cargar_HappyNonAdmin = BuildOk(selectedId, logs)
    Exit Function
EH:
    Test_ComercialHelper_Seleccionar_Cargar_HappyNonAdmin = BuildFail(p_Error, logs)
End Function

' 10. Seleccionar_Cargar — sad empty selection
Public Function Test_ComercialHelper_Seleccionar_Cargar_SadEmptySelection() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubComercialesDict(Array("FIRST"))

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_Seleccionar_Cargar("", col, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("enableEditar") <> False Then Err.Raise 1001, , "expected enableEditar=false"
    If Not parsed("payload")("entity") Is Nothing Then Err.Raise 1002, , "expected entity=null"

    Test_ComercialHelper_Seleccionar_Cargar_SadEmptySelection = BuildOk("empty", logs)
    Exit Function
EH:
    Test_ComercialHelper_Seleccionar_Cargar_SadEmptySelection = BuildFail(p_Error, logs)
End Function

' 11. Eliminar_Borrar — sad cancelled
Public Function Test_ComercialHelper_Eliminar_Borrar_SadCancelled() As String
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
    Set entity = BuildStubComercial(seedId, "CANCEL-TEST", "desc")

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_Eliminar_Borrar(entity, vbNo, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim countAfter As Long
    countAfter = DCount("*", "TbComerciales", "IDComercial=" & seedId)
    If countAfter <> 1 Then Err.Raise 1002, , "row should still exist after cancel"

    Test_ComercialHelper_Eliminar_Borrar_SadCancelled = BuildOk("cancelled", logs)
    Call TeardownFixture
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    On Error GoTo 0
    Test_ComercialHelper_Eliminar_Borrar_SadCancelled = BuildFail(p_Error, logs)
End Function

' 12. Eliminar_Borrar — happy deleted
Public Function Test_ComercialHelper_Eliminar_Borrar_HappyDeleted() As String
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
    countBefore = DCount("*", "TbComerciales", "IDComercial=" & seedId)
    If countBefore <> 1 Then Err.Raise 1002, , "expected 1 row before delete"

    Dim entity As Object
    Set entity = BuildStubComercial(seedId, "DELETE-TEST", "desc")

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_Eliminar_Borrar(entity, vbYes, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim countAfter As Long
    countAfter = DCount("*", "TbComerciales", "IDComercial=" & seedId)
    If countAfter <> 0 Then Err.Raise 1003, , "expected 0 rows after delete, got " & countAfter

    Test_ComercialHelper_Eliminar_Borrar_HappyDeleted = BuildOk("deleted", logs)
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    On Error GoTo 0
    Test_ComercialHelper_Eliminar_Borrar_HappyDeleted = BuildFail(p_Error, logs)
End Function

' 13. Eliminar_Borrar — sad no entity
Public Function Test_ComercialHelper_Eliminar_Borrar_SadNoEntity() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_Eliminar_Borrar(Nothing, vbYes, p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false"

    Test_ComercialHelper_Eliminar_Borrar_SadNoEntity = BuildOk("no-entity-rejected", logs)
    Exit Function
EH:
    Test_ComercialHelper_Eliminar_Borrar_SadNoEntity = BuildFail(p_Error, logs)
End Function

' 14. DobleClick_AbrirEdicion — happy choose
Public Function Test_ComercialHelper_DobleClick_AbrirEdicion_HappyChoose() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_DobleClick_AbrirEdicion(True, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("action")) <> "choose" Then Err.Raise 1001, , "expected action=choose"

    Test_ComercialHelper_DobleClick_AbrirEdicion_HappyChoose = BuildOk("choose", logs)
    Exit Function
EH:
    Test_ComercialHelper_DobleClick_AbrirEdicion_HappyChoose = BuildFail(p_Error, logs)
End Function

' 15. DobleClick_AbrirEdicion — happy edit
Public Function Test_ComercialHelper_DobleClick_AbrirEdicion_HappyEdit() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_DobleClick_AbrirEdicion(False, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("action")) <> "edit" Then Err.Raise 1001, , "expected action=edit"

    Test_ComercialHelper_DobleClick_AbrirEdicion_HappyEdit = BuildOk("edit", logs)
    Exit Function
EH:
    Test_ComercialHelper_DobleClick_AbrirEdicion_HappyEdit = BuildFail(p_Error, logs)
End Function

' 16. DobleClick_AbrirEdicion — happy none
Public Function Test_ComercialHelper_DobleClick_AbrirEdicion_HappyNone() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_DobleClick_AbrirEdicion(False, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("action")) <> "none" Then Err.Raise 1001, , "expected action=none"

    Test_ComercialHelper_DobleClick_AbrirEdicion_HappyNone = BuildOk("none", logs)
    Exit Function
EH:
    Test_ComercialHelper_DobleClick_AbrirEdicion_HappyNone = BuildFail(p_Error, logs)
End Function

' 17. DobleClick_AbrirEdicion — edge choose takes precedence
Public Function Test_ComercialHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modComercialHelper.Comercial_DobleClick_AbrirEdicion(True, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("action")) <> "choose" Then Err.Raise 1001, , "expected choose precedence"

    Test_ComercialHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence = BuildOk("choose-precedence", logs)
    Exit Function
EH:
    Test_ComercialHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence = BuildFail(p_Error, logs)
End Function

' 18. RunAll — wrapper for Dysflow manifest discovery
Public Function Test_ComercialHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_ComercialHelper_Abrir_Inicializar_HappyAdminNoArgs", _
        "Test_ComercialHelper_Abrir_Inicializar_HappyNonAdminNoArgs", _
        "Test_ComercialHelper_Abrir_Inicializar_HappyAdminWithArgs", _
        "Test_ComercialHelper_Buscar_Listar_HappyAllRows", _
        "Test_ComercialHelper_Buscar_Listar_EdgeEmptyCollection", _
        "Test_ComercialHelper_Buscar_Listar_HappyFilteredRows", _
        "Test_ComercialHelper_Buscar_Listar_AdversarialFilterWithSemicolons", _
        "Test_ComercialHelper_Seleccionar_Cargar_HappyAdmin", _
        "Test_ComercialHelper_Seleccionar_Cargar_HappyNonAdmin", _
        "Test_ComercialHelper_Seleccionar_Cargar_SadEmptySelection", _
        "Test_ComercialHelper_Eliminar_Borrar_SadCancelled", _
        "Test_ComercialHelper_Eliminar_Borrar_HappyDeleted", _
        "Test_ComercialHelper_Eliminar_Borrar_SadNoEntity", _
        "Test_ComercialHelper_DobleClick_AbrirEdicion_HappyChoose", _
        "Test_ComercialHelper_DobleClick_AbrirEdicion_HappyEdit", _
        "Test_ComercialHelper_DobleClick_AbrirEdicion_HappyNone", _
        "Test_ComercialHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence" _
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
        Test_ComercialHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_ComercialHelper_RunAll = BuildOk("all-passed", logs)
    End If
    On Error Resume Next
    Call TeardownFixture
    On Error GoTo 0
    Exit Function
EH:
    On Error Resume Next
    Dim p_Error As String
    p_Error = "Test_ComercialHelper_RunAll EH: " & Err.Description
    Call TeardownFixture
    On Error GoTo 0
    Test_ComercialHelper_RunAll = BuildFail(p_Error, logs)
End Function
