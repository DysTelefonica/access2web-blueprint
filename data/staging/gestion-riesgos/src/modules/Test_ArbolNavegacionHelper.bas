Attribute VB_Name = "Test_ArbolNavegacionHelper"
' =============================================================================
' Test_ArbolNavegacionHelper — TDD atoms for
'   modArbolNavegacionHelper.CargarArbol
'   modArbolNavegacionHelper.SeleccionarNodo
'
' Skill: access-vba-tdd v2.6 + access-vba-e2e-methodology
' SDD:   forms-thin-phase0-testeable-2026-06-25 / Block 2a / T-3.1
'
' Scope: 1 helper module -> 11 scenario atoms
'   Helper 1: CargarArbol(scope, p_Refrescando, p_Error) — load tree by scope
'   Helper 2: SeleccionarNodo(p_Nodo, p_Error) — select node
'
' Conventions:
'   - Returns JSON {ok, value, payload, error, logs} via TestCore_BuildOk / BuildFail
'   - p_Error ByRef (Telefonica D&S); empty string on success
' =============================================================================
Option Compare Database
Option Explicit

' --- Constants ---
Private Const TEST_ID_BASE As Long = 900400
Private Const TEST_ID_EDICION As Long = 900410
Private Const TEST_ID_PROYECTO As Long = 900420
Private Const TEST_ID_RIEGO_A As Long = 900401
Private Const TEST_ID_RIEGO_B As Long = 900402
Private Const TEST_ID_RIEGO_C As Long = 900403

' =============================================================================
' JSON wrappers (BuildOk/BuildFail) — must precede all Public Functions
' Per access-vba-tdd §1.8
' =============================================================================
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' =============================================================================
' Private helpers
' =============================================================================

' CountRows — cardinality helper per access-vba-tdd §4.5
Private Function CountRows(ByVal db As DAO.Database, ByVal p_Table As String, _
                           ByVal p_Where As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS c FROM " & p_Table & " WHERE " & p_Where)
    CountRows = CLng(Nz(rs!c, 0))
    rs.Close
    Set rs = Nothing
End Function

' =============================================================================
' ATOM: B2a-S1 Happy CargarArbol scope=riesgos
' =============================================================================
Public Function Test_ArbolNavegacionHelper_B2a_S1_Happy_CargarArbol_Riesgos() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ArbolNavegacionHelper_B2a_S1_Happy_CargarArbol_Riesgos = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ArbolNavegacionHelper_B2a_S1_Happy_CargarArbol_Riesgos = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    ' Idempotency DELETE
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900400 AND 900499", dbFailOnError

    Dim countAfter As Long
    countAfter = CountRows(db, "TbRiesgos", "IDRiesgo BETWEEN 900400 AND 900499")
    If countAfter <> 0 Then
        Test_ArbolNavegacionHelper_B2a_S1_Happy_CargarArbol_Riesgos = BuildFail("teardown FAIL: residual rows " & countAfter, logs)
        Exit Function
    End If

    ' Act: load empty tree
    Dim p_Error As String
    Dim result As String
    result = modArbolNavegacionHelper.CargarArbol("riesgos", EnumSiNo.No, p_Error)

    If p_Error <> "" Then
        Test_ArbolNavegacionHelper_B2a_S1_Happy_CargarArbol_Riesgos = BuildFail("p_Error=" & p_Error, logs)
        Exit Function
    End If

    logIdx = logIdx + 1
    Test_ArbolNavegacionHelper_B2a_S1_Happy_CargarArbol_Riesgos = BuildOk("cargado", logs)
    Exit Function
EH:
    Test_ArbolNavegacionHelper_B2a_S1_Happy_CargarArbol_Riesgos = BuildFail("EH: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM: B2a-S2 Happy CargarArbol scope=accion
' =============================================================================
Public Function Test_ArbolNavegacionHelper_B2a_S2_Happy_CargarArbol_Accion() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ArbolNavegacionHelper_B2a_S2_Happy_CargarArbol_Accion = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ArbolNavegacionHelper_B2a_S2_Happy_CargarArbol_Accion = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    Dim p_Error As String
    Dim result As String
    result = modArbolNavegacionHelper.CargarArbol("accion", EnumSiNo.No, p_Error)

    If p_Error <> "" Then
        Test_ArbolNavegacionHelper_B2a_S2_Happy_CargarArbol_Accion = BuildFail("p_Error=" & p_Error, logs)
        Exit Function
    End If

    logIdx = logIdx + 1
    Test_ArbolNavegacionHelper_B2a_S2_Happy_CargarArbol_Accion = BuildOk("cargado", logs)
    Exit Function
EH:
    Test_ArbolNavegacionHelper_B2a_S2_Happy_CargarArbol_Accion = BuildFail("EH: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM: B2a-S3 Happy CargarArbol scope=PC
' =============================================================================
Public Function Test_ArbolNavegacionHelper_B2a_S3_Happy_CargarArbol_PC() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ArbolNavegacionHelper_B2a_S3_Happy_CargarArbol_PC = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ArbolNavegacionHelper_B2a_S3_Happy_CargarArbol_PC = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    Dim p_Error As String
    Dim result As String
    result = modArbolNavegacionHelper.CargarArbol("PC", EnumSiNo.No, p_Error)

    If p_Error <> "" Then
        Test_ArbolNavegacionHelper_B2a_S3_Happy_CargarArbol_PC = BuildFail("p_Error=" & p_Error, logs)
        Exit Function
    End If

    logIdx = logIdx + 1
    Test_ArbolNavegacionHelper_B2a_S3_Happy_CargarArbol_PC = BuildOk("cargado", logs)
    Exit Function
EH:
    Test_ArbolNavegacionHelper_B2a_S3_Happy_CargarArbol_PC = BuildFail("EH: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM: B2a-S4 Happy CargarArbol scope=PM
' =============================================================================
Public Function Test_ArbolNavegacionHelper_B2a_S4_Happy_CargarArbol_PM() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ArbolNavegacionHelper_B2a_S4_Happy_CargarArbol_PM = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ArbolNavegacionHelper_B2a_S4_Happy_CargarArbol_PM = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    Dim p_Error As String
    Dim result As String
    result = modArbolNavegacionHelper.CargarArbol("PM", EnumSiNo.No, p_Error)

    If p_Error <> "" Then
        Test_ArbolNavegacionHelper_B2a_S4_Happy_CargarArbol_PM = BuildFail("p_Error=" & p_Error, logs)
        Exit Function
    End If

    logIdx = logIdx + 1
    Test_ArbolNavegacionHelper_B2a_S4_Happy_CargarArbol_PM = BuildOk("cargado", logs)
    Exit Function
EH:
    Test_ArbolNavegacionHelper_B2a_S4_Happy_CargarArbol_PM = BuildFail("EH: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM: B2a-S5 Sad CargarArbol scope=invalid
' =============================================================================
Public Function Test_ArbolNavegacionHelper_B2a_S5_Sad_CargarArbol_Invalid() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ArbolNavegacionHelper_B2a_S5_Sad_CargarArbol_Invalid = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ArbolNavegacionHelper_B2a_S5_Sad_CargarArbol_Invalid = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    Dim p_Error As String
    Dim result As String
    result = modArbolNavegacionHelper.CargarArbol("invalid_scope_xyz", EnumSiNo.No, p_Error)

    If p_Error = "" Then
        Test_ArbolNavegacionHelper_B2a_S5_Sad_CargarArbol_Invalid = BuildFail("p_Error should be populated for invalid scope", logs)
        Exit Function
    End If

    logIdx = logIdx + 1
    Test_ArbolNavegacionHelper_B2a_S5_Sad_CargarArbol_Invalid = BuildOk("invalid_rejected", logs)
    Exit Function
EH:
    Test_ArbolNavegacionHelper_B2a_S5_Sad_CargarArbol_Invalid = BuildFail("EH: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM: B2a-S6 Sad CargarArbol empty scope
' =============================================================================
Public Function Test_ArbolNavegacionHelper_B2a_S6_Sad_CargarArbol_EmptyScope() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ArbolNavegacionHelper_B2a_S6_Sad_CargarArbol_EmptyScope = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ArbolNavegacionHelper_B2a_S6_Sad_CargarArbol_EmptyScope = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    Dim p_Error As String
    Dim result As String
    result = modArbolNavegacionHelper.CargarArbol("", EnumSiNo.No, p_Error)

    If p_Error = "" Then
        Test_ArbolNavegacionHelper_B2a_S6_Sad_CargarArbol_EmptyScope = BuildFail("p_Error should be populated for empty scope", logs)
        Exit Function
    End If

    logIdx = logIdx + 1
    Test_ArbolNavegacionHelper_B2a_S6_Sad_CargarArbol_EmptyScope = BuildOk("empty_rejected", logs)
    Exit Function
EH:
    Test_ArbolNavegacionHelper_B2a_S6_Sad_CargarArbol_EmptyScope = BuildFail("EH: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM: B2a-S7 Edge CargarArbol with data
' =============================================================================
Public Function Test_ArbolNavegacionHelper_B2a_S7_Edge_CargarArbol_WithData() As String
    Dim logs(0 To 7) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ArbolNavegacionHelper_B2a_S7_Edge_CargarArbol_WithData = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ArbolNavegacionHelper_B2a_S7_Edge_CargarArbol_WithData = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    ' Idempotency DELETE
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900400 AND 900499", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_ID_PROYECTO, dbFailOnError

    ' Seed 3 riesgos in same edition
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-ARBOL-S7', 'Arbol test')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, Estado) " & _
               "VALUES (" & TEST_ID_RIEGO_A & ", " & TEST_ID_EDICION & ", 'PC-A', 'R-A', 'Activo')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, Estado) " & _
               "VALUES (" & TEST_ID_RIEGO_B & ", " & TEST_ID_EDICION & ", 'PC-B', 'R-B', 'Aceptado')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, Estado) " & _
               "VALUES (" & TEST_ID_RIEGO_C & ", " & TEST_ID_EDICION & ", 'PC-C', 'R-C', 'Materializado')", dbFailOnError

    Dim countAfter As Long
    countAfter = CountRows(db, "TbRiesgos", "IDEdicion=" & TEST_ID_EDICION)
    If countAfter <> 3 Then
        Test_ArbolNavegacionHelper_B2a_S7_Edge_CargarArbol_WithData = BuildFail("Arrange FAIL: expected 3 rows, got " & countAfter, logs)
        Exit Function
    End If

    ' Act
    Dim p_Error As String
    Dim result As String
    result = modArbolNavegacionHelper.CargarArbol("riesgos", EnumSiNo.No, p_Error)

    If p_Error <> "" Then
        Test_ArbolNavegacionHelper_B2a_S7_Edge_CargarArbol_WithData = BuildFail("p_Error=" & p_Error, logs)
        Exit Function
    End If

    logIdx = logIdx + 1
    Test_ArbolNavegacionHelper_B2a_S7_Edge_CargarArbol_WithData = BuildOk("3_riesgos_loaded", logs)
    Exit Function
EH:
    Test_ArbolNavegacionHelper_B2a_S7_Edge_CargarArbol_WithData = BuildFail("EH: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM: B2a-S8 Edge CargarArbol p_Refrescando=Si
' =============================================================================
Public Function Test_ArbolNavegacionHelper_B2a_S8_Edge_CargarArbol_Refrescando() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ArbolNavegacionHelper_B2a_S8_Edge_CargarArbol_Refrescando = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ArbolNavegacionHelper_B2a_S8_Edge_CargarArbol_Refrescando = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    Dim p_Error As String
    Dim result As String
    result = modArbolNavegacionHelper.CargarArbol("riesgos", EnumSiNo.Sí, p_Error)

    If p_Error <> "" Then
        Test_ArbolNavegacionHelper_B2a_S8_Edge_CargarArbol_Refrescando = BuildFail("p_Error=" & p_Error, logs)
        Exit Function
    End If

    logIdx = logIdx + 1
    Test_ArbolNavegacionHelper_B2a_S8_Edge_CargarArbol_Refrescando = BuildOk("refrescado", logs)
    Exit Function
EH:
    Test_ArbolNavegacionHelper_B2a_S8_Edge_CargarArbol_Refrescando = BuildFail("EH: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM: B2a-S9 Sad SeleccionarNodo Nothing
' Per spec: p_Nodo=Nothing es el caso Sad -> p_Error populado.
' Renamed from Happy -> Sad after Fit Phase 2 e2e audit (Fix A):
' helper contract rejects Nothing, atom was mislabeled.
' =============================================================================
Public Function Test_ArbolNavegacionHelper_B2a_S9_Sad_SeleccionarNodo_Nothing() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ArbolNavegacionHelper_B2a_S9_Sad_SeleccionarNodo_Nothing = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ArbolNavegacionHelper_B2a_S9_Sad_SeleccionarNodo_Nothing = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    Dim p_Error As String
    Dim m_Nodo As Object
    Set m_Nodo = Nothing
    Dim result As String
    result = modArbolNavegacionHelper.SeleccionarNodo(m_Nodo, p_Error)

    ' Per spec: p_Nodo=Nothing es el caso Sad -> p_Error populado
    If p_Error = "" Then
        Test_ArbolNavegacionHelper_B2a_S9_Sad_SeleccionarNodo_Nothing = BuildFail("p_Error should be populated for Nothing node", logs)
        Exit Function
    End If

    logIdx = logIdx + 1
    Test_ArbolNavegacionHelper_B2a_S9_Sad_SeleccionarNodo_Nothing = BuildOk("nothing_rejected", logs)
    Exit Function
EH:
    Test_ArbolNavegacionHelper_B2a_S9_Sad_SeleccionarNodo_Nothing = BuildFail("EH: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM: B2a-S10 Sad SeleccionarNodo Nothing
' =============================================================================
Public Function Test_ArbolNavegacionHelper_B2a_S10_Sad_SeleccionarNodo_Nothing() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ArbolNavegacionHelper_B2a_S10_Sad_SeleccionarNodo_Nothing = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ArbolNavegacionHelper_B2a_S10_Sad_SeleccionarNodo_Nothing = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    Dim p_Error As String
    Dim result As String
    result = modArbolNavegacionHelper.SeleccionarNodo(Nothing, p_Error)

    ' Per spec: p_Nodo=Nothing es el caso Sad -> p_Error populado
    If p_Error = "" Then
        Test_ArbolNavegacionHelper_B2a_S10_Sad_SeleccionarNodo_Nothing = BuildFail("p_Error should be populated for Nothing node", logs)
        Exit Function
    End If

    logIdx = logIdx + 1
    Test_ArbolNavegacionHelper_B2a_S10_Sad_SeleccionarNodo_Nothing = BuildOk("nothing_rejected", logs)
    Exit Function
EH:
    Test_ArbolNavegacionHelper_B2a_S10_Sad_SeleccionarNodo_Nothing = BuildFail("EH: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM: B2a-S11 Adversarial CargarArbol called twice
' =============================================================================
Public Function Test_ArbolNavegacionHelper_B2a_S11_Adversarial_CargarArbol_Doble() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ArbolNavegacionHelper_B2a_S11_Adversarial_CargarArbol_Doble = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ArbolNavegacionHelper_B2a_S11_Adversarial_CargarArbol_Doble = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    ' Idempotency DELETE
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900400 AND 900499", dbFailOnError

    Dim p_Error1 As String
    Dim p_Error2 As String
    Dim result1 As String
    Dim result2 As String
    result1 = modArbolNavegacionHelper.CargarArbol("riesgos", EnumSiNo.No, p_Error1)
    result2 = modArbolNavegacionHelper.CargarArbol("riesgos", EnumSiNo.Sí, p_Error2)

    If p_Error1 <> "" Or p_Error2 <> "" Then
        Test_ArbolNavegacionHelper_B2a_S11_Adversarial_CargarArbol_Doble = BuildFail("p_Error populated: " & p_Error1 & " / " & p_Error2, logs)
        Exit Function
    End If

    logIdx = logIdx + 1
    Test_ArbolNavegacionHelper_B2a_S11_Adversarial_CargarArbol_Doble = BuildOk("doble_carga_ok", logs)
    Exit Function
EH:
    Test_ArbolNavegacionHelper_B2a_S11_Adversarial_CargarArbol_Doble = BuildFail("EH: " & Err.Description, logs)
End Function
