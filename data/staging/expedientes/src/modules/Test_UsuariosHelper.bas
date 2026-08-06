Attribute VB_Name = "Test_UsuariosHelper"
Option Compare Database
Option Explicit

' Test_UsuariosHelper — REWORK (2026-06-26)
' Pure-data TDD atoms for modUsuariosHelper.bas
' (Form_FormUsuariosGestion SELECT-only, see docs/audit/usuarios-gestion-thin.md).
'
' Anti-pattern removed (was in PR #31 commit c86d460):
'   - Test_UsuariosHelper_OpenForm(ByRef p_Form, ByRef p_Error) used to call
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
'
' Convention:
'   - Atoms use ONLY Global public names (access-vba-tdd §1.1.1).
'   - Each atom returns JSON: {"ok":true|false,"value":...,"payload":null,"error":...,"logs":[...]}.
'   - Atoms never call MsgBox or pop up UI.

' === Module-level constants (all at top per vba-access §10.1) =======================

' Fixture ID base — entities created by atoms use IDs in this range.
Private Const TEST_BASE_ID As Long = 900720


' === Local helpers (all at top per vba-access §10.1) ================================

Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

' --- BuildStubUsuario ----------------------------------------------------------
' Builds a Scripting.Dictionary {ID, Nombre, FechaBaja}. FechaBaja may be Null
' (active), a date (inactive), or a string. The helper treats any non-date as active.
Private Function BuildStubUsuario( _
    ByVal p_ID As String, _
    ByVal p_Nombre As String, _
    ByVal p_FechaBaja As Variant _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ID") = p_ID
    d("Nombre") = p_Nombre
    d("FechaBaja") = p_FechaBaja
    Set BuildStubUsuario = d
End Function

' --- BuildStubUsuariosDict ----------------------------------------------------
' Builds a Dictionary {CStr(ID) -> stubEntity} from an array of {Nombre, FechaBaja} tuples.
' IDs are TEST_BASE_ID, TEST_BASE_ID+1, ...
Private Function BuildStubUsuariosDict(ByVal p_Items As Variant) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(p_Items) To UBound(p_Items)
        Dim idStr As String
        idStr = CStr(TEST_BASE_ID + i)
        Set d(idStr) = BuildStubUsuario(idStr, _
                                       CStr(p_Items(i)(0)), _
                                       CStr(p_Items(i)(1)))
    Next i
    Set BuildStubUsuariosDict = d
End Function


' === Public atoms ====================================================================

' ---------------------------------------------------------------------------
' 1. Abrir_Inicializar — happy (admin=true, no OpenArgs)
'    Expected: hasElegir=false, hasOpenArgs=false
' ---------------------------------------------------------------------------
Public Function Test_UsuariosHelper_Abrir_Inicializar_HappyAdminNoArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modUsuariosHelper.Usuarios_Abrir_Inicializar( _
        True, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    If InStr(json, """ok"":true") = 0 Then Err.Raise 1001, , "expected ok=true, got " & json

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hasElegir") <> False Then Err.Raise 1002, , "expected hasElegir=false"
    If parsed("payload")("hasOpenArgs") <> False Then Err.Raise 1003, , "expected hasOpenArgs=false"

    Test_UsuariosHelper_Abrir_Inicializar_HappyAdminNoArgs = BuildOk("admin-no-args", logs)
    Exit Function
EH:
    Test_UsuariosHelper_Abrir_Inicializar_HappyAdminNoArgs = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 2. Abrir_Inicializar — happy (admin=true, has OpenArgs)
'    Expected: hasElegir=true, hasOpenArgs=true
' ---------------------------------------------------------------------------
Public Function Test_UsuariosHelper_Abrir_Inicializar_HappyAdminWithArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modUsuariosHelper.Usuarios_Abrir_Inicializar( _
        True, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hasElegir") <> True Then Err.Raise 1002, , "expected hasElegir=true"
    If parsed("payload")("hasOpenArgs") <> True Then Err.Raise 1003, , "expected hasOpenArgs=true"

    Test_UsuariosHelper_Abrir_Inicializar_HappyAdminWithArgs = BuildOk("admin-with-args", logs)
    Exit Function
EH:
    Test_UsuariosHelper_Abrir_Inicializar_HappyAdminWithArgs = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 3. Buscar_Listar — happy (Todos -> 3 rows including inactive)
' ---------------------------------------------------------------------------
Public Function Test_UsuariosHelper_Buscar_Listar_HappyTodos() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    ' {Nombre, FechaBaja}; FechaBaja="" = active, FechaBaja=date string = inactive.
    Set col = BuildStubUsuariosDict(Array( _
        Array("User1-Active", ""), _
        Array("User2-Active", ""), _
        Array("User3-Inactive", "2025-01-01")))

    Dim p_Error As String
    Dim json As String
    json = modUsuariosHelper.Usuarios_Buscar_Listar(col, "", "Todos", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 3 Then
        Err.Raise 1002, , "expected count=3 (Todos), got " & parsed("payload")("count")
    End If

    Test_UsuariosHelper_Buscar_Listar_HappyTodos = BuildOk(3, logs)
    Exit Function
EH:
    Test_UsuariosHelper_Buscar_Listar_HappyTodos = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 4. Buscar_Listar — happy (Sí -> only active, 2 rows)
' ---------------------------------------------------------------------------
Public Function Test_UsuariosHelper_Buscar_Listar_HappyActiveOnly() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubUsuariosDict(Array( _
        Array("Active1", ""), _
        Array("Inactive1", "2025-01-01"), _
        Array("Active2", "")))

    Dim p_Error As String
    Dim json As String
    json = modUsuariosHelper.Usuarios_Buscar_Listar(col, "", "Sí", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 2 Then
        Err.Raise 1002, , "expected count=2 (active only), got " & parsed("payload")("count")
    End If

    Test_UsuariosHelper_Buscar_Listar_HappyActiveOnly = BuildOk(2, logs)
    Exit Function
EH:
    Test_UsuariosHelper_Buscar_Listar_HappyActiveOnly = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 5. Buscar_Listar — happy (No -> only inactive, 2 rows)
' ---------------------------------------------------------------------------
Public Function Test_UsuariosHelper_Buscar_Listar_HappyInactiveOnly() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubUsuariosDict(Array( _
        Array("Active1", ""), _
        Array("Inactive1", "2025-01-01"), _
        Array("Inactive2", "2025-02-01")))

    Dim p_Error As String
    Dim json As String
    json = modUsuariosHelper.Usuarios_Buscar_Listar(col, "", "No", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 2 Then
        Err.Raise 1002, , "expected count=2 (inactive only), got " & parsed("payload")("count")
    End If

    Test_UsuariosHelper_Buscar_Listar_HappyInactiveOnly = BuildOk(2, logs)
    Exit Function
EH:
    Test_UsuariosHelper_Buscar_Listar_HappyInactiveOnly = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 6. Buscar_Listar — edge (empty collection -> 0 rows)
' ---------------------------------------------------------------------------
Public Function Test_UsuariosHelper_Buscar_Listar_EdgeEmptyCollection() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modUsuariosHelper.Usuarios_Buscar_Listar(Nothing, "", "Todos", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 0 Then
        Err.Raise 1002, , "expected count=0"
    End If

    Test_UsuariosHelper_Buscar_Listar_EdgeEmptyCollection = BuildOk(0, logs)
    Exit Function
EH:
    Test_UsuariosHelper_Buscar_Listar_EdgeEmptyCollection = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 7. Buscar_Listar — happy (text filter on Nombre, Todos -> 2 rows)
' ---------------------------------------------------------------------------
Public Function Test_UsuariosHelper_Buscar_Listar_HappyTextFilter() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubUsuariosDict(Array( _
        Array("AlphaUser", ""), _
        Array("BetaUser", ""), _
        Array("AlphaAdmin", "")))

    Dim p_Error As String
    Dim json As String
    json = modUsuariosHelper.Usuarios_Buscar_Listar(col, "Alpha", "Todos", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 2 Then
        Err.Raise 1002, , "expected count=2 (Alpha text filter), got " & parsed("payload")("count")
    End If

    Dim rs As String
    rs = CStr(parsed("payload")("rowSource"))
    If InStr(rs, "BetaUser") > 0 Then
        Err.Raise 1003, , "rowSource should NOT contain BetaUser"
    End If

    Test_UsuariosHelper_Buscar_Listar_HappyTextFilter = BuildOk(2, logs)
    Exit Function
EH:
    Test_UsuariosHelper_Buscar_Listar_HappyTextFilter = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 8. Buscar_Listar — edge (256-char filter that does not match -> 0 rows)
' ---------------------------------------------------------------------------
Public Function Test_UsuariosHelper_Buscar_Listar_EdgeLongFilter() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubUsuariosDict(Array( _
        Array("SHORT", "")))

    Dim longFilter As String
    longFilter = String(256, "X")

    Dim p_Error As String
    Dim json As String
    json = modUsuariosHelper.Usuarios_Buscar_Listar(col, longFilter, "Todos", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CLng(parsed("payload")("count")) <> 0 Then
        Err.Raise 1002, , "expected count=0 for 256-char non-matching filter"
    End If

    Test_UsuariosHelper_Buscar_Listar_EdgeLongFilter = BuildOk(0, logs)
    Exit Function
EH:
    Test_UsuariosHelper_Buscar_Listar_EdgeLongFilter = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 9. Seleccionar_Cargar — happy (selection present, admin)
'    Expected: entity=stub, hasSelection=true
' ---------------------------------------------------------------------------
Public Function Test_UsuariosHelper_Seleccionar_Cargar_HappyAdmin() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubUsuariosDict(Array( _
        Array("FIRST", "")))

    Dim selectedId As String
    selectedId = CStr(TEST_BASE_ID + 1)

    Dim p_Error As String
    Dim json As String
    json = modUsuariosHelper.Usuarios_Seleccionar_Cargar( _
        selectedId, col, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hasSelection") <> True Then
        Err.Raise 1002, , "expected hasSelection=true"
    End If
    If parsed("payload")("entity") Is Nothing Then
        Err.Raise 1003, , "expected entity populated"
    End If
    If CStr(parsed("payload")("entity")("ID")) <> selectedId Then
        Err.Raise 1004, , "expected ID=" & selectedId & ", got " & parsed("payload")("entity")("ID")
    End If
    If CStr(parsed("payload")("entity")("Nombre")) <> "FIRST" Then
        Err.Raise 1005, , "expected Nombre=FIRST, got " & parsed("payload")("entity")("Nombre")
    End If

    Test_UsuariosHelper_Seleccionar_Cargar_HappyAdmin = BuildOk(selectedId, logs)
    Exit Function
EH:
    Test_UsuariosHelper_Seleccionar_Cargar_HappyAdmin = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 10. Seleccionar_Cargar — sad (empty selection -> entity=null)
' ---------------------------------------------------------------------------
Public Function Test_UsuariosHelper_Seleccionar_Cargar_SadEmptySelection() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubUsuariosDict(Array( _
        Array("FIRST", "")))

    Dim p_Error As String
    Dim json As String
    json = modUsuariosHelper.Usuarios_Seleccionar_Cargar( _
        "", col, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If parsed("payload")("hasSelection") <> False Then
        Err.Raise 1002, , "expected hasSelection=false"
    End If
    If Not parsed("payload")("entity") Is Nothing Then
        Err.Raise 1003, , "expected entity=null"
    End If

    Test_UsuariosHelper_Seleccionar_Cargar_SadEmptySelection = BuildOk("empty", logs)
    Exit Function
EH:
    Test_UsuariosHelper_Seleccionar_Cargar_SadEmptySelection = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 11. Limpiar_Reset — happy (returns activosDefault="Todos")
' ---------------------------------------------------------------------------
Public Function Test_UsuariosHelper_Limpiar_Reset_HappyCleared() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modUsuariosHelper.Usuarios_Limpiar_Reset(p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("activosDefault")) <> "Todos" Then
        Err.Raise 1002, , "expected activosDefault=Todos, got " & parsed("payload")("activosDefault")
    End If

    Test_UsuariosHelper_Limpiar_Reset_HappyCleared = BuildOk("cleared", logs)
    Exit Function
EH:
    Test_UsuariosHelper_Limpiar_Reset_HappyCleared = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 12. DobleClick_AbrirDetalle — happy (choose dispatched when cmdElegir visible)
' ---------------------------------------------------------------------------
Public Function Test_UsuariosHelper_DobleClick_AbrirDetalle_HappyChoose() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modUsuariosHelper.Usuarios_DobleClick_AbrirDetalle( _
        True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("action")) <> "choose" Then
        Err.Raise 1002, , "expected action=choose, got " & parsed("payload")("action")
    End If

    Test_UsuariosHelper_DobleClick_AbrirDetalle_HappyChoose = BuildOk("choose", logs)
    Exit Function
EH:
    Test_UsuariosHelper_DobleClick_AbrirDetalle_HappyChoose = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 13. DobleClick_AbrirDetalle — happy (none when cmdElegir hidden — SELECT-only)
' ---------------------------------------------------------------------------
Public Function Test_UsuariosHelper_DobleClick_AbrirDetalle_HappyNone() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modUsuariosHelper.Usuarios_DobleClick_AbrirDetalle( _
        False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)

    If CStr(parsed("payload")("action")) <> "none" Then
        Err.Raise 1002, , "expected action=none (SELECT-only), got " & parsed("payload")("action")
    End If

    Test_UsuariosHelper_DobleClick_AbrirDetalle_HappyNone = BuildOk("none", logs)
    Exit Function
EH:
    Test_UsuariosHelper_DobleClick_AbrirDetalle_HappyNone = BuildFail(p_Error, logs)
End Function

' ---------------------------------------------------------------------------
' 14. RunAll — wrapper for Dysflow manifest discovery (access-vba-tdd §1.1.1)
' ---------------------------------------------------------------------------
Public Function Test_UsuariosHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_UsuariosHelper_Abrir_Inicializar_HappyAdminNoArgs", _
        "Test_UsuariosHelper_Abrir_Inicializar_HappyAdminWithArgs", _
        "Test_UsuariosHelper_Buscar_Listar_HappyTodos", _
        "Test_UsuariosHelper_Buscar_Listar_HappyActiveOnly", _
        "Test_UsuariosHelper_Buscar_Listar_HappyInactiveOnly", _
        "Test_UsuariosHelper_Buscar_Listar_EdgeEmptyCollection", _
        "Test_UsuariosHelper_Buscar_Listar_HappyTextFilter", _
        "Test_UsuariosHelper_Buscar_Listar_EdgeLongFilter", _
        "Test_UsuariosHelper_Seleccionar_Cargar_HappyAdmin", _
        "Test_UsuariosHelper_Seleccionar_Cargar_SadEmptySelection", _
        "Test_UsuariosHelper_Limpiar_Reset_HappyCleared", _
        "Test_UsuariosHelper_DobleClick_AbrirDetalle_HappyChoose", _
        "Test_UsuariosHelper_DobleClick_AbrirDetalle_HappyNone" _
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
        Test_UsuariosHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_UsuariosHelper_RunAll = BuildOk("all-passed", logs)
    End If
    Exit Function
EH:
    Dim p_Error As String
    p_Error = "Test_UsuariosHelper_RunAll EH: " & Err.Description
    Test_UsuariosHelper_RunAll = BuildFail(p_Error, logs)
End Function
