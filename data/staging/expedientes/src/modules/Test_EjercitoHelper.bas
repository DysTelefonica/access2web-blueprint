Attribute VB_Name = "Test_EjercitoHelper"
Option Compare Database
Option Explicit

' Test_EjercitoHelper — REWORK TDD atoms for modEjercitoHelper.bas
' (Form_FormEjercitosGestion, see docs/audit/ejercito-gestion-pure-data.md).

' === Module-level constants (all at top per vba-access §10.1) =======================

Private Const TEST_BASE_ID As Long = 900520

Private Const TEST_FIELD_SEP As String = ";"


' === Local helpers (all at top per vba-access §10.1) ================================

Private Function BuildOk(ByVal p_Value As Variant, ByRef p_Logs() As String) As String
    BuildOk = TestingCore_BuildOk(p_Value, p_Logs)
End Function

Private Function BuildFail(ByVal p_Error As String, ByRef p_Logs() As String) As String
    BuildFail = TestingCore_BuildFail(p_Error, p_Logs)
End Function

Private Function BuildStubEjercito( _
    ByVal p_ID As String, _
    ByVal p_Nombre As String, _
    ByVal p_Desc As String _
) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("IDEjercito") = p_ID
    d("Ejercito") = p_Nombre
    d("DESCRIPCION") = p_Desc
    Set BuildStubEjercito = d
End Function

Private Function BuildStubEjercitosDict(ByVal p_Names As Variant) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(p_Names) To UBound(p_Names)
        Dim idStr As String
        idStr = CStr(TEST_BASE_ID + i)
        Set d(idStr) = BuildStubEjercito(idStr, CStr(p_Names(i)), "Desc " & CStr(p_Names(i)))
    Next i
    Set BuildStubEjercitosDict = d
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
    sql = "INSERT INTO TbEjercitos (IDEjercito, Ejercito, Descripcion) VALUES (" & _
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
    before = DCount("*", "TbEjercitos", "IDEjercito >= " & TEST_BASE_ID)
    db.Execute "DELETE FROM TbEjercitos WHERE IDEjercito >= " & TEST_BASE_ID, dbFailOnError
    On Error GoTo 0
    TeardownFixture = before
End Function


' === Public atoms ====================================================================

Public Function Test_EjercitoHelper_Abrir_Inicializar_HappyAdminNoArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modEjercitoHelper.Ejercito_Abrir_Inicializar(True, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("showAlta") <> True Then Err.Raise 1001, , "expected showAlta=true"
    If parsed("payload")("showElegir") <> False Then Err.Raise 1002, , "expected showElegir=false"

    Test_EjercitoHelper_Abrir_Inicializar_HappyAdminNoArgs = BuildOk("admin-no-args", logs)
    Exit Function
EH:
    Test_EjercitoHelper_Abrir_Inicializar_HappyAdminNoArgs = BuildFail(p_Error, logs)
End Function

Public Function Test_EjercitoHelper_Abrir_Inicializar_HappyNonAdminNoArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modEjercitoHelper.Ejercito_Abrir_Inicializar(False, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("showAlta") <> False Then Err.Raise 1001, , "expected showAlta=false"

    Test_EjercitoHelper_Abrir_Inicializar_HappyNonAdminNoArgs = BuildOk("non-admin-no-args", logs)
    Exit Function
EH:
    Test_EjercitoHelper_Abrir_Inicializar_HappyNonAdminNoArgs = BuildFail(p_Error, logs)
End Function

Public Function Test_EjercitoHelper_Abrir_Inicializar_HappyAdminWithArgs() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modEjercitoHelper.Ejercito_Abrir_Inicializar(True, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("showElegir") <> True Then Err.Raise 1001, , "expected showElegir=true"

    Test_EjercitoHelper_Abrir_Inicializar_HappyAdminWithArgs = BuildOk("admin-with-args", logs)
    Exit Function
EH:
    Test_EjercitoHelper_Abrir_Inicializar_HappyAdminWithArgs = BuildFail(p_Error, logs)
End Function

Public Function Test_EjercitoHelper_Buscar_Listar_HappyAllRows() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubEjercitosDict(Array("ALPHA", "BETA", "GAMMA"))

    Dim p_Error As String
    Dim json As String
    json = modEjercitoHelper.Ejercito_Buscar_Listar(col, "", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CLng(parsed("payload")("count")) <> 3 Then Err.Raise 1001, , "expected count=3"

    Dim rs As String
    rs = CStr(parsed("payload")("rowSource"))
    ' 3-col listbox includes DESCRIPCION.
    If InStr(rs, "Desc ALPHA") = 0 Then Err.Raise 1002, , "rowSource missing Desc ALPHA"

    Test_EjercitoHelper_Buscar_Listar_HappyAllRows = BuildOk(3, logs)
    Exit Function
EH:
    Test_EjercitoHelper_Buscar_Listar_HappyAllRows = BuildFail(p_Error, logs)
End Function

Public Function Test_EjercitoHelper_Buscar_Listar_EdgeEmptyCollection() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modEjercitoHelper.Ejercito_Buscar_Listar(Nothing, "", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CLng(parsed("payload")("count")) <> 0 Then Err.Raise 1001, , "expected count=0"

    Test_EjercitoHelper_Buscar_Listar_EdgeEmptyCollection = BuildOk(0, logs)
    Exit Function
EH:
    Test_EjercitoHelper_Buscar_Listar_EdgeEmptyCollection = BuildFail(p_Error, logs)
End Function

Public Function Test_EjercitoHelper_Buscar_Listar_HappyFilteredRows() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubEjercitosDict(Array("ALPHA-A", "BETA", "ALPHA-B"))

    Dim p_Error As String
    Dim json As String
    json = modEjercitoHelper.Ejercito_Buscar_Listar(col, "ALPHA", p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CLng(parsed("payload")("count")) <> 2 Then Err.Raise 1001, , "expected count=2 (ALPHA match)"

    Test_EjercitoHelper_Buscar_Listar_HappyFilteredRows = BuildOk(2, logs)
    Exit Function
EH:
    Test_EjercitoHelper_Buscar_Listar_HappyFilteredRows = BuildFail(p_Error, logs)
End Function

Public Function Test_EjercitoHelper_Seleccionar_Cargar_HappyAdmin() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubEjercitosDict(Array("FIRST"))

    Dim selectedId As String
    selectedId = CStr(TEST_BASE_ID + 1)

    Dim p_Error As String
    Dim json As String
    json = modEjercitoHelper.Ejercito_Seleccionar_Cargar(selectedId, col, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("enableEditar") <> True Then Err.Raise 1001, , "expected enableEditar=true"
    If parsed("payload")("enableEliminar") <> True Then Err.Raise 1002, , "expected enableEliminar=true for admin"

    Test_EjercitoHelper_Seleccionar_Cargar_HappyAdmin = BuildOk(selectedId, logs)
    Exit Function
EH:
    Test_EjercitoHelper_Seleccionar_Cargar_HappyAdmin = BuildFail(p_Error, logs)
End Function

Public Function Test_EjercitoHelper_Seleccionar_Cargar_SadEmptySelection() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(3)
    On Error GoTo EH

    Dim col As Object
    Set col = BuildStubEjercitosDict(Array("FIRST"))

    Dim p_Error As String
    Dim json As String
    json = modEjercitoHelper.Ejercito_Seleccionar_Cargar("", col, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If parsed("payload")("enableEditar") <> False Then Err.Raise 1001, , "expected enableEditar=false"

    Test_EjercitoHelper_Seleccionar_Cargar_SadEmptySelection = BuildOk("empty", logs)
    Exit Function
EH:
    Test_EjercitoHelper_Seleccionar_Cargar_SadEmptySelection = BuildFail(p_Error, logs)
End Function

Public Function Test_EjercitoHelper_Eliminar_Borrar_SadCancelled() As String
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
    Set entity = BuildStubEjercito(seedId, "CANCEL-TEST", "desc")

    Dim p_Error As String
    Dim json As String
    json = modEjercitoHelper.Ejercito_Eliminar_Borrar(entity, vbNo, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim countAfter As Long
    countAfter = DCount("*", "TbEjercitos", "IDEjercito=" & seedId)
    If countAfter <> 1 Then Err.Raise 1002, , "row should still exist after cancel"

    Test_EjercitoHelper_Eliminar_Borrar_SadCancelled = BuildOk("cancelled", logs)
    Call TeardownFixture
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    On Error GoTo 0
    Test_EjercitoHelper_Eliminar_Borrar_SadCancelled = BuildFail(p_Error, logs)
End Function

Public Function Test_EjercitoHelper_Eliminar_Borrar_HappyDeleted() As String
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
    countBefore = DCount("*", "TbEjercitos", "IDEjercito=" & seedId)
    If countBefore <> 1 Then Err.Raise 1002, , "expected 1 row before delete"

    Dim entity As Object
    Set entity = BuildStubEjercito(seedId, "DELETE-TEST", "desc")

    Dim p_Error As String
    Dim json As String
    json = modEjercitoHelper.Ejercito_Eliminar_Borrar(entity, vbYes, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim countAfter As Long
    countAfter = DCount("*", "TbEjercitos", "IDEjercito=" & seedId)
    If countAfter <> 0 Then Err.Raise 1003, , "expected 0 rows after delete, got " & countAfter

    Test_EjercitoHelper_Eliminar_Borrar_HappyDeleted = BuildOk("deleted", logs)
    Exit Function
EH:
    On Error Resume Next
    Call TeardownFixture
    On Error GoTo 0
    Test_EjercitoHelper_Eliminar_Borrar_HappyDeleted = BuildFail(p_Error, logs)
End Function

Public Function Test_EjercitoHelper_Eliminar_Borrar_SadNoEntity() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modEjercitoHelper.Ejercito_Eliminar_Borrar(Nothing, vbYes, p_Error)
    If InStr(json, """ok"":false") = 0 Then Err.Raise 1001, , "expected ok=false"

    Test_EjercitoHelper_Eliminar_Borrar_SadNoEntity = BuildOk("no-entity-rejected", logs)
    Exit Function
EH:
    Test_EjercitoHelper_Eliminar_Borrar_SadNoEntity = BuildFail(p_Error, logs)
End Function

Public Function Test_EjercitoHelper_DobleClick_AbrirEdicion_HappyChoose() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modEjercitoHelper.Ejercito_DobleClick_AbrirEdicion(True, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("action")) <> "choose" Then Err.Raise 1001, , "expected action=choose"

    Test_EjercitoHelper_DobleClick_AbrirEdicion_HappyChoose = BuildOk("choose", logs)
    Exit Function
EH:
    Test_EjercitoHelper_DobleClick_AbrirEdicion_HappyChoose = BuildFail(p_Error, logs)
End Function

Public Function Test_EjercitoHelper_DobleClick_AbrirEdicion_HappyEdit() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modEjercitoHelper.Ejercito_DobleClick_AbrirEdicion(False, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("action")) <> "edit" Then Err.Raise 1001, , "expected action=edit"

    Test_EjercitoHelper_DobleClick_AbrirEdicion_HappyEdit = BuildOk("edit", logs)
    Exit Function
EH:
    Test_EjercitoHelper_DobleClick_AbrirEdicion_HappyEdit = BuildFail(p_Error, logs)
End Function

Public Function Test_EjercitoHelper_DobleClick_AbrirEdicion_HappyNone() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modEjercitoHelper.Ejercito_DobleClick_AbrirEdicion(False, False, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("action")) <> "none" Then Err.Raise 1001, , "expected action=none"

    Test_EjercitoHelper_DobleClick_AbrirEdicion_HappyNone = BuildOk("none", logs)
    Exit Function
EH:
    Test_EjercitoHelper_DobleClick_AbrirEdicion_HappyNone = BuildFail(p_Error, logs)
End Function

Public Function Test_EjercitoHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim p_Error As String
    Dim json As String
    json = modEjercitoHelper.Ejercito_DobleClick_AbrirEdicion(True, True, p_Error)
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(json)
    If CStr(parsed("payload")("action")) <> "choose" Then Err.Raise 1001, , "expected choose precedence"

    Test_EjercitoHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence = BuildOk("choose-precedence", logs)
    Exit Function
EH:
    Test_EjercitoHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence = BuildFail(p_Error, logs)
End Function

Public Function Test_EjercitoHelper_RunAll() As String
    Dim logs() As String
    logs = TestingCore_InitLogs(2)
    On Error GoTo EH

    Dim atoms As Variant
    atoms = Array( _
        "Test_EjercitoHelper_Abrir_Inicializar_HappyAdminNoArgs", _
        "Test_EjercitoHelper_Abrir_Inicializar_HappyNonAdminNoArgs", _
        "Test_EjercitoHelper_Abrir_Inicializar_HappyAdminWithArgs", _
        "Test_EjercitoHelper_Buscar_Listar_HappyAllRows", _
        "Test_EjercitoHelper_Buscar_Listar_EdgeEmptyCollection", _
        "Test_EjercitoHelper_Buscar_Listar_HappyFilteredRows", _
        "Test_EjercitoHelper_Seleccionar_Cargar_HappyAdmin", _
        "Test_EjercitoHelper_Seleccionar_Cargar_SadEmptySelection", _
        "Test_EjercitoHelper_Eliminar_Borrar_SadCancelled", _
        "Test_EjercitoHelper_Eliminar_Borrar_HappyDeleted", _
        "Test_EjercitoHelper_Eliminar_Borrar_SadNoEntity", _
        "Test_EjercitoHelper_DobleClick_AbrirEdicion_HappyChoose", _
        "Test_EjercitoHelper_DobleClick_AbrirEdicion_HappyEdit", _
        "Test_EjercitoHelper_DobleClick_AbrirEdicion_HappyNone", _
        "Test_EjercitoHelper_DobleClick_AbrirEdicion_EdgeChooseTakesPrecedence" _
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
        Test_EjercitoHelper_RunAll = BuildFail("firstFailure=" & firstFailure, logs)
    Else
        Test_EjercitoHelper_RunAll = BuildOk("all-passed", logs)
    End If
    On Error Resume Next
    Call TeardownFixture
    On Error GoTo 0
    Exit Function
EH:
    On Error Resume Next
    Dim p_Error As String
    p_Error = "Test_EjercitoHelper_RunAll EH: " & Err.Description
    Call TeardownFixture
    On Error GoTo 0
    Test_EjercitoHelper_RunAll = BuildFail(p_Error, logs)
End Function
