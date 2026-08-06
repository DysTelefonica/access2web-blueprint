Attribute VB_Name = "Test_RiesgoMaterializacion"
Option Compare Database
Option Explicit

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

' ============================================================
' F.8 — Public desmaterialization/reversal path (issue #34, PR 8)
' MaterializacionQuitarRegistrar clears TbRiesgos.FechaMaterializado
' and appends one 'No' row to TbRiesgosMaterializaciones, preserving
' the previous 'Sí' row and the unrelated guard row.
' ============================================================
Public Function Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha() As String
    On Error GoTo EH

    Dim logs(0 To 12) As String
    Dim errMsg As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim riesgo As riesgo
    Dim codigoRiesgo As String
    Dim rowsForKeyBefore As Long
    Dim rowsForKeyAfter As Long
    Dim rowsPreviousSiBefore As Long
    Dim rowsPreviousSiAfter As Long
    Dim rowsNoBefore As Long
    Dim rowsNoAfter As Long
    Dim rowsGuardBefore As Long
    Dim rowsGuardAfter As Long
    Dim previousFechaBefore As Date
    Dim previousEstadoBefore As String
    Dim guardFechaBefore As Date
    Dim guardEstadoBefore As String
    Dim guardParaNCBefore As String
    Dim guardFechaDecisionBefore As Date
    Dim guardIDNCBefore As Long
    Dim riesgoEstadoAfter As String
    Dim noRowEstado As String
    Dim noRowFecha As Date
    Dim cleanupDb As DAO.Database
    Dim cleanupErr As String
    Const PREV_MAT_ID As Long = 910970
    Const GUARD_MAT_ID As Long = 910972

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: SeedAll parent graph (project=900501, edition=900502, risk=900503)"
    logs(2) = "3. Arrange: resolve fixture CodigoRiesgo without SELECT TOP 1"
    logs(3) = "4. Arrange: set TbRiesgos FechaMaterializado=#06/01/2026# and Estado='Materializado'"
    logs(4) = "5. Arrange: previous 'Sí' materialization row ID=" & PREV_MAT_ID
    logs(5) = "6. Arrange: unrelated guard row ID=" & GUARD_MAT_ID
    logs(6) = "7. Assert pre-state: FechaMaterializado, Estado='Materializado', previous='Sí', no 'No' rows"
    logs(7) = "8. Act: Constructor.getRiesgo(...).MaterializacionQuitarRegistrar"
    logs(8) = "9. Assert: FechaMaterializado is Null and Estado recalculates away from Materializado"
    logs(9) = "10. Assert: exactly one new 'No' history row with Estado equal to post-reversal risk state"
    logs(10) = "11. Assert: previous 'Sí' row remains unchanged"
    logs(11) = "12. Assert: unrelated guard row remains unchanged"
    logs(12) = "13. Teardown: delete deterministic materialization rows + TeardownAll"

    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail(errMsg, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT CodigoRiesgo FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Fixture riesgo " & Test_Fixtures.Cache_RiesgoId & " not found", logs)
        GoTo Teardown
    End If
    codigoRiesgo = CStr(Nz(rs.fields("CodigoRiesgo").value, ""))
    rs.Close: Set rs = Nothing
    If codigoRiesgo = "" Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Fixture riesgo CodigoRiesgo is empty", logs)
        GoTo Teardown
    End If

    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE IDProyecto=" & _
        Test_Fixtures.Cache_ProyectoId & " AND CodigoRiesgo='" & _
        Replace(codigoRiesgo, "'", "''") & "' AND ID >= " & PREV_MAT_ID, dbFailOnError
    Err.Clear
    On Error GoTo EH

    previousFechaBefore = DateSerial(2026, 6, 1)
    previousEstadoBefore = "Materializado"
    guardFechaBefore = DateSerial(2026, 6, 3)
    guardEstadoBefore = "Materializado"
    guardParaNCBefore = "No"
    guardFechaDecisionBefore = DateSerial(2026, 6, 4)
    guardIDNCBefore = 880972

    db.Execute "UPDATE TbRiesgos SET FechaMaterializado=#06/01/2026#, " & _
        "Estado='Materializado' WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, dbFailOnError

    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, Estado) " & _
        "VALUES (" & PREV_MAT_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "#06/01/2026#, 'Sí', '" & previousEstadoBefore & "')", dbFailOnError

    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, " & _
        "Estado, IDNC, ParaNC, FechaDecison) " & _
        "VALUES (" & GUARD_MAT_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "#06/03/2026#, 'Sí', '" & guardEstadoBefore & "', " & _
        guardIDNCBefore & ", '" & guardParaNCBefore & "', #06/04/2026#)", dbFailOnError

    rowsForKeyBefore = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId & " AND CodigoRiesgo='" & _
        Replace(codigoRiesgo, "'", "''") & "'")
    rowsPreviousSiBefore = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & PREV_MAT_ID & " AND EsMaterializacion='Sí'")
    rowsNoBefore = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId & " AND CodigoRiesgo='" & _
        Replace(codigoRiesgo, "'", "''") & "' AND EsMaterializacion='No'")
    rowsGuardBefore = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones WHERE ID=" & GUARD_MAT_ID)

    If rowsForKeyBefore <> 2 Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Initial key cardinality: expected 2, got " & rowsForKeyBefore, logs)
        GoTo Teardown
    End If
    If rowsPreviousSiBefore <> 1 Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Previous 'Sí' row cardinality: expected 1, got " & rowsPreviousSiBefore, logs)
        GoTo Teardown
    End If
    If rowsNoBefore <> 0 Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Initial 'No' rows should be 0, got " & rowsNoBefore, logs)
        GoTo Teardown
    End If
    If rowsGuardBefore <> 1 Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Initial guard cardinality: expected 1, got " & rowsGuardBefore, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT FechaMaterializado, Estado FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Fixture riesgo disappeared before Act", logs)
        GoTo Teardown
    End If
    If IsNull(rs.fields("FechaMaterializado").value) Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Pre-state FechaMaterializado should be #06/01/2026#, got Null", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("FechaMaterializado").value) <> previousFechaBefore Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Pre-state FechaMaterializado changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("Estado").value, "")), "Materializado", vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Pre-state Estado should be 'Materializado'", logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    errMsg = ""
    Set riesgo = Constructor.getRiesgo(CStr(Test_Fixtures.Cache_RiesgoId), , , errMsg)
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Constructor.getRiesgo failed: " & errMsg, logs)
        GoTo Teardown
    End If
    If riesgo Is Nothing Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Constructor.getRiesgo returned Nothing", logs)
        GoTo Teardown
    End If

    errMsg = ""
    riesgo.MaterializacionQuitarRegistrar EnumRiesgoEstado.Materializado, errMsg
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("MaterializacionQuitarRegistrar should not fail: " & errMsg, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT FechaMaterializado, Estado FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Fixture riesgo disappeared after Act", logs)
        GoTo Teardown
    End If
    If Not IsNull(rs.fields("FechaMaterializado").value) Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("FechaMaterializado should be Null after reversal", logs)
        GoTo Teardown
    End If
    riesgoEstadoAfter = CStr(Nz(rs.fields("Estado").value, ""))
    If StrComp(riesgoEstadoAfter, "Materializado", vbTextCompare) = 0 Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Estado should recalculate away from Materializado", logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    rowsForKeyAfter = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId & " AND CodigoRiesgo='" & _
        Replace(codigoRiesgo, "'", "''") & "'")
    rowsNoAfter = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId & " AND CodigoRiesgo='" & _
        Replace(codigoRiesgo, "'", "''") & "' AND EsMaterializacion='No'")
    rowsPreviousSiAfter = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & PREV_MAT_ID & " AND EsMaterializacion='Sí'")
    rowsGuardAfter = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones WHERE ID=" & GUARD_MAT_ID)

    If rowsForKeyAfter <> rowsForKeyBefore + 1 Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Final key cardinality: expected " & (rowsForKeyBefore + 1) & ", got " & rowsForKeyAfter, logs)
        GoTo Teardown
    End If
    If rowsNoAfter <> 1 Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Expected exactly one new 'No' row, got " & rowsNoAfter, logs)
        GoTo Teardown
    End If
    If rowsPreviousSiAfter <> rowsPreviousSiBefore Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Previous 'Sí' row cardinality changed", logs)
        GoTo Teardown
    End If
    If rowsGuardAfter <> rowsGuardBefore Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Guard cardinality changed", logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT Fecha, Estado FROM TbRiesgosMaterializaciones " & _
        "WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId & _
        " AND CodigoRiesgo='" & Replace(codigoRiesgo, "'", "''") & "'" & _
        " AND EsMaterializacion='No'", dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("No desmaterialization row found", logs)
        GoTo Teardown
    End If
    noRowEstado = CStr(Nz(rs.fields("Estado").value, ""))
    If StrComp(noRowEstado, riesgoEstadoAfter, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("'No' row Estado should equal risk Estado '" & riesgoEstadoAfter & "', got '" & noRowEstado & "'", logs)
        GoTo Teardown
    End If
    If IsNull(rs.fields("Fecha").value) Or Not IsDate(rs.fields("Fecha").value) Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("'No' row Fecha should be a valid Date", logs)
        GoTo Teardown
    End If
    noRowFecha = CDate(rs.fields("Fecha").value)
    If noRowFecha <> Date Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("'No' row Fecha should equal today", logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    Set rs = db.OpenRecordset( _
        "SELECT Fecha, Estado FROM TbRiesgosMaterializaciones WHERE ID=" & PREV_MAT_ID, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Previous 'Sí' row disappeared", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("Fecha").value) <> previousFechaBefore Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Previous 'Sí' Fecha changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("Estado").value, "")), previousEstadoBefore, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Previous 'Sí' Estado changed", logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    Set rs = db.OpenRecordset( _
        "SELECT Fecha, Estado, IDNC, ParaNC, FechaDecison " & _
        "FROM TbRiesgosMaterializaciones WHERE ID=" & GUARD_MAT_ID, dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
            BuildFail("Guard row disappeared", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("Fecha").value) <> guardFechaBefore Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = BuildFail("Guard Fecha changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("Estado").value, "")), guardEstadoBefore, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = BuildFail("Guard Estado changed", logs)
        GoTo Teardown
    End If
    If CLng(Nz(rs.fields("IDNC").value, 0)) <> guardIDNCBefore Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = BuildFail("Guard IDNC changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("ParaNC").value, "")), guardParaNCBefore, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = BuildFail("Guard ParaNC changed", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("FechaDecison").value) <> guardFechaDecisionBefore Then
        Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = BuildFail("Guard FechaDecison changed", logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
        BuildOk("materializacion_quitar_registrar_desmaterializa", logs)

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set riesgo = Nothing
    Set cleanupDb = Test_Fixtures.GetTestDb(cleanupErr)
    If Not cleanupDb Is Nothing Then
        cleanupDb.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE IDProyecto=" & _
            Test_Fixtures.Cache_ProyectoId & " AND CodigoRiesgo='" & _
            Replace(codigoRiesgo, "'", "''") & "' AND ID >= " & PREV_MAT_ID
        Set cleanupDb = Nothing
    End If
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
    Exit Function

EH:
    Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha = _
        BuildFail("Test_RiesgoMaterializacion_MaterializacionQuitarRegistrar_CreaDesmaterializacionYLimpiaFecha: " & _
            Err.description, logs)
    Resume Teardown
End Function

' ============================================================
' F.7 — Eliminar happy/safety path (issue #34, PR 7)
' Deletes only the targeted materialization row. The fixture
' creates one target and one unrelated guard row under the
' deterministic SeedAll parent graph; Eliminar must reduce target
' cardinality 1 -> 0, preserve the guard row, preserve parent rows,
' and return an empty errMsg.
' ============================================================
Public Function Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo() As String
    On Error GoTo EH

    Dim logs(0 To 9) As String
    Dim errMsg As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim mat As RiesgoMaterializacion
    Dim codigoRiesgo As String
    Dim rowsTargetBefore As Long
    Dim rowsTargetAfter As Long
    Dim rowsGuardBefore As Long
    Dim rowsGuardAfter As Long
    Dim parentProyectoBefore As Long
    Dim parentEdicionBefore As Long
    Dim parentRiesgoBefore As Long
    Dim parentProyectoAfter As Long
    Dim parentEdicionAfter As Long
    Dim parentRiesgoAfter As Long
    Dim guardFechaBefore As Date
    Dim guardEstadoBefore As String
    Dim guardParaNCBefore As String
    Dim guardFechaDecisionBefore As Date
    Dim guardIDNCBefore As Long
    Dim cleanupDb As DAO.Database
    Dim cleanupErr As String
    Const TARGET_MAT_ID As Long = 910960
    Const GUARD_MAT_ID As Long = 910962

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: SeedAll parent graph (project=900501, edition=900502, risk=900503)"
    logs(2) = "3. Arrange: resolve fixture CodigoRiesgo without SELECT TOP 1"
    logs(3) = "4. Arrange: target row ID=" & TARGET_MAT_ID & " and guard row ID=" & GUARD_MAT_ID
    logs(4) = "5. Arrange: assert target=1, guard=1, parent graph present"
    logs(5) = "6. Act: Constructor.getRiesgoMaterializado(target).Eliminar errMsg"
    logs(6) = "7. Assert: errMsg='' and target row count 1 -> 0"
    logs(7) = "8. Assert: guard row remains unchanged"
    logs(8) = "9. Assert: parent project, edition, and risk rows remain present"
    logs(9) = "10. Teardown: delete deterministic materialization rows + TeardownAll"

    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = BuildFail(errMsg, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT CodigoRiesgo FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = _
            BuildFail("Fixture riesgo " & Test_Fixtures.Cache_RiesgoId & " not found", logs)
        GoTo Teardown
    End If
    codigoRiesgo = CStr(Nz(rs.fields("CodigoRiesgo").value, ""))
    rs.Close: Set rs = Nothing
    If codigoRiesgo = "" Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = _
            BuildFail("Fixture riesgo CodigoRiesgo is empty", logs)
        GoTo Teardown
    End If

    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID IN (" & _
        TARGET_MAT_ID & ", " & GUARD_MAT_ID & ")", dbFailOnError
    Err.Clear
    On Error GoTo EH

    guardFechaBefore = DateSerial(2026, 6, 15)
    guardEstadoBefore = "Materializado"
    guardParaNCBefore = "No"
    guardFechaDecisionBefore = DateSerial(2026, 6, 16)
    guardIDNCBefore = 880962

    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, Estado) " & _
        "VALUES (" & TARGET_MAT_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "#06/10/2026#, 'Sí', 'Materializado')", dbFailOnError

    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, " & _
        "Estado, IDNC, ParaNC, FechaDecison) " & _
        "VALUES (" & GUARD_MAT_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "#06/15/2026#, 'Sí', '" & guardEstadoBefore & "', " & _
        guardIDNCBefore & ", '" & guardParaNCBefore & "', #06/16/2026#)", dbFailOnError

    rowsTargetBefore = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones WHERE ID=" & TARGET_MAT_ID)
    rowsGuardBefore = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones WHERE ID=" & GUARD_MAT_ID)
    parentProyectoBefore = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbProyectos WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId)
    parentEdicionBefore = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbProyectosEdiciones WHERE IDEdicion=" & Test_Fixtures.Cache_EdicionId)
    parentRiesgoBefore = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId)

    If rowsTargetBefore <> 1 Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = _
            BuildFail("Initial target cardinality: expected 1, got " & rowsTargetBefore, logs)
        GoTo Teardown
    End If
    If rowsGuardBefore <> 1 Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = _
            BuildFail("Initial guard cardinality: expected 1, got " & rowsGuardBefore, logs)
        GoTo Teardown
    End If
    If parentProyectoBefore <> 1 Or parentEdicionBefore <> 1 Or parentRiesgoBefore <> 1 Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = _
            BuildFail("Parent graph missing before Act", logs)
        GoTo Teardown
    End If

    errMsg = ""
    Set mat = Constructor.getRiesgoMaterializado(CStr(TARGET_MAT_ID), errMsg)
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = _
            BuildFail("getRiesgoMaterializado failed: " & errMsg, logs)
        GoTo Teardown
    End If
    If mat Is Nothing Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = _
            BuildFail("getRiesgoMaterializado returned Nothing for ID=" & TARGET_MAT_ID, logs)
        GoTo Teardown
    End If

    errMsg = ""
    mat.Eliminar errMsg
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = _
            BuildFail("Eliminar should not fail: " & errMsg, logs)
        GoTo Teardown
    End If

    rowsTargetAfter = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones WHERE ID=" & TARGET_MAT_ID)
    rowsGuardAfter = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones WHERE ID=" & GUARD_MAT_ID)
    If rowsTargetAfter <> 0 Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = _
            BuildFail("Target row should be deleted: expected 0, got " & rowsTargetAfter, logs)
        GoTo Teardown
    End If
    If rowsGuardAfter <> rowsGuardBefore Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = _
            BuildFail("Guard cardinality changed: expected " & rowsGuardBefore & ", got " & rowsGuardAfter, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT Fecha, Estado, IDNC, ParaNC, FechaDecison " & _
        "FROM TbRiesgosMaterializaciones WHERE ID=" & GUARD_MAT_ID, dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = _
            BuildFail("Guard row disappeared", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("Fecha").value) <> guardFechaBefore Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = BuildFail("Guard Fecha changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("Estado").value, "")), guardEstadoBefore, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = BuildFail("Guard Estado changed", logs)
        GoTo Teardown
    End If
    If CLng(Nz(rs.fields("IDNC").value, 0)) <> guardIDNCBefore Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = BuildFail("Guard IDNC changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("ParaNC").value, "")), guardParaNCBefore, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = BuildFail("Guard ParaNC changed", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("FechaDecison").value) <> guardFechaDecisionBefore Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = BuildFail("Guard FechaDecison changed", logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    parentProyectoAfter = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbProyectos WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId)
    parentEdicionAfter = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbProyectosEdiciones WHERE IDEdicion=" & Test_Fixtures.Cache_EdicionId)
    parentRiesgoAfter = CountRowsBySql(db, "SELECT COUNT(*) AS Cnt FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId)
    If parentProyectoAfter <> parentProyectoBefore Or _
       parentEdicionAfter <> parentEdicionBefore Or _
       parentRiesgoAfter <> parentRiesgoBefore Then
        Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = _
            BuildFail("Parent graph cardinality changed after Eliminar", logs)
        GoTo Teardown
    End If

    Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = _
        BuildOk("eliminar_borra_solo_materializacion_objetivo", logs)

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set mat = Nothing
    Set cleanupDb = Test_Fixtures.GetTestDb(cleanupErr)
    If Not cleanupDb Is Nothing Then
        cleanupDb.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID IN (" & _
            TARGET_MAT_ID & ", " & GUARD_MAT_ID & ")"
        Set cleanupDb = Nothing
    End If
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
    Exit Function

EH:
    Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo = _
        BuildFail("Test_RiesgoMaterializacion_Eliminar_BorraSoloMaterializacionObjetivo: " & _
            Err.description, logs)
    Resume Teardown
End Function

Private Function CountRowsBySql(ByVal db As DAO.Database, ByVal sql As String) As Long
    Dim rs As DAO.Recordset

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Then
        CountRowsBySql = 0
    Else
        CountRowsBySql = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close
    Set rs = Nothing
End Function

' ============================================================
' Issue #34 — RegistrarParaNONC happy path
' Persists ParaNC='No', keeps IDNC Null, stores the explicit
' FechaDecison, and preserves unrelated lifecycle fields.
' ============================================================
Public Function Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC() As String
    On Error GoTo EH

    Dim logs(0 To 9) As String
    Dim errMsg As String
    Dim fechaDecision As Date
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim mat As RiesgoMaterializacion
    Dim codigoRiesgo As String
    Dim rowsForKeyBefore As Long
    Dim rowsForKeyAfter As Long
    Dim fechaOriginal As Date
    Dim estadoOriginal As String
    Dim cleanupDb As DAO.Database
    Dim cleanupErr As String
    Const PREV_MAT_ID As Long = 910900

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: SeedAll fixture graph (project=900501, riesgo=900503)"
    logs(2) = "3. Arrange: resolve fixture CodigoRiesgo without SELECT TOP 1"
    logs(3) = "4. Arrange: pre-existing 'Sí' materialization row at ID=" & PREV_MAT_ID & " with Null NC decision fields"
    logs(4) = "5. Arrange: count fixture key before RegistrarParaNONC"
    logs(5) = "6. Act: RegistrarParaNONC(#01/15/2026#)"
    logs(6) = "7. Assert: same row persists ParaNC='No', IDNC Null, FechaDecison=#01/15/2026#"
    logs(7) = "8. Assert: cardinalidad rowsForKey unchanged (1 -> 1)"
    logs(8) = "9. Assert: unrelated lifecycle fields unchanged"
    logs(9) = "10. Teardown: delete pre-existing materialization row + TeardownAll"

    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = BuildFail(errMsg, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT CodigoRiesgo FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = _
            BuildFail("Fixture riesgo " & Test_Fixtures.Cache_RiesgoId & " not found", logs)
        GoTo Teardown
    End If
    codigoRiesgo = CStr(Nz(rs.fields("CodigoRiesgo").value, ""))
    rs.Close: Set rs = Nothing
    If codigoRiesgo = "" Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = _
            BuildFail("Fixture riesgo CodigoRiesgo is empty", logs)
        GoTo Teardown
    End If

    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID=" & PREV_MAT_ID, dbFailOnError
    Err.Clear
    On Error GoTo EH

    fechaOriginal = DateSerial(2026, 1, 10)
    estadoOriginal = "Materializado"
    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, Estado) " & _
        "VALUES (" & PREV_MAT_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "#01/10/2026#, 'Sí', '" & estadoOriginal & "')", dbFailOnError

    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId & _
        " AND CodigoRiesgo='" & Replace(codigoRiesgo, "'", "''") & "'", _
        dbOpenSnapshot)
    If rs.EOF Then
        rowsForKeyBefore = 0
    Else
        rowsForKeyBefore = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close: Set rs = Nothing
    If rowsForKeyBefore <> 1 Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = _
            BuildFail("Cardinalidad inicial: expected exactly 1 materialization row for fixture key, got " & _
                rowsForKeyBefore, logs)
        GoTo Teardown
    End If

    errMsg = ""
    Set mat = Constructor.getRiesgoMaterializado(CStr(PREV_MAT_ID), errMsg)
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = _
            BuildFail("getRiesgoMaterializado failed: " & errMsg, logs)
        GoTo Teardown
    End If
    If mat Is Nothing Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = _
            BuildFail("getRiesgoMaterializado returned Nothing for ID=" & PREV_MAT_ID, logs)
        GoTo Teardown
    End If

    fechaDecision = DateSerial(2026, 1, 15)
    errMsg = ""
    mat.RegistrarParaNONC CStr(fechaDecision), errMsg
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = _
            BuildFail("RegistrarParaNONC should not fail: " & errMsg, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, " & _
        "Estado, IDNC, ParaNC, FechaDecison FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & PREV_MAT_ID, dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = _
            BuildFail("Materialization row not found for ID=" & PREV_MAT_ID, logs)
        GoTo Teardown
    End If

    If CLng(rs.fields("ID").value) <> PREV_MAT_ID Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = BuildFail("ID changed after RegistrarParaNONC", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("ParaNC").value, "")), "No", vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = _
            BuildFail("ParaNC on row should be 'No', got '" & Nz(rs.fields("ParaNC").value, "") & "'", logs)
        GoTo Teardown
    End If
    If Not IsNull(rs.fields("IDNC").value) Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = _
            BuildFail("IDNC on row should be Null, got " & rs.fields("IDNC").value, logs)
        GoTo Teardown
    End If
    If IsNull(rs.fields("FechaDecison").value) Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = _
            BuildFail("FechaDecison on row should be #01/15/2026#, got Null", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("FechaDecison").value) <> fechaDecision Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = _
            BuildFail("FechaDecison on row should be #01/15/2026#, got " & rs.fields("FechaDecison").value, logs)
        GoTo Teardown
    End If

    If CLng(rs.fields("IDProyecto").value) <> Test_Fixtures.Cache_ProyectoId Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = BuildFail("IDProyecto changed", logs)
        GoTo Teardown
    End If
    If CLng(rs.fields("IDEdicion").value) <> Test_Fixtures.Cache_EdicionId Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = BuildFail("IDEdicion changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("CodigoRiesgo").value, "")), codigoRiesgo, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = BuildFail("CodigoRiesgo changed", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("Fecha").value) <> fechaOriginal Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = BuildFail("Fecha changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("EsMaterializacion").value, "")), "Sí", vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = BuildFail("EsMaterializacion changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("Estado").value, "")), estadoOriginal, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = BuildFail("Estado changed", logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId & _
        " AND CodigoRiesgo='" & Replace(codigoRiesgo, "'", "''") & "'", _
        dbOpenSnapshot)
    If rs.EOF Then
        rowsForKeyAfter = 0
    Else
        rowsForKeyAfter = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close: Set rs = Nothing
    If rowsForKeyAfter <> 1 Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = _
            BuildFail("Cardinalidad final: expected exactly 1 materialization row for fixture key, got " & _
                rowsForKeyAfter, logs)
        GoTo Teardown
    End If
    If rowsForKeyAfter <> rowsForKeyBefore Then
        Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = _
            BuildFail("Cardinalidad: expected rowsForKey unchanged (" & rowsForKeyBefore & _
                "), got " & rowsForKeyAfter, logs)
        GoTo Teardown
    End If

    Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = _
        BuildOk("registrar_para_nonc_persiste_decision_no_nc", logs)

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set mat = Nothing
    Set cleanupDb = Test_Fixtures.GetTestDb(cleanupErr)
    If Not cleanupDb Is Nothing Then
        cleanupDb.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID=" & PREV_MAT_ID
        Set cleanupDb = Nothing
    End If
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
    Exit Function

EH:
    Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC = _
        BuildFail("Test_RiesgoMaterializacion_RegistrarParaNONC_PersisteDecisionNoNC: " & _
            Err.description, logs)
    Resume Teardown
End Function

Public Function Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha() As String
    On Error GoTo EH

    Dim logs(0 To 10) As String
    Dim errMsg As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim mat As RiesgoMaterializacion
    Dim codigoRiesgo As String
    Dim rowsForTargetBefore As Long
    Dim rowsForTargetAfter As Long
    Dim originalFecha As Date
    Dim newFecha As Date
    Dim targetEstado As String
    Dim unrelatedFechaBefore As Date
    Dim unrelatedEstadoBefore As String
    Dim unrelatedParaNCBefore As String
    Dim unrelatedFechaDecisionBefore As Date
    Dim cleanupDb As DAO.Database
    Dim cleanupErr As String
    Const TARGET_MAT_ID As Long = 910910
    Const UNRELATED_MAT_ID As Long = 910912

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: SeedAll fixture graph (project=900501, riesgo=900503)"
    logs(2) = "3. Arrange: resolve fixture CodigoRiesgo without SELECT TOP 1"
    logs(3) = "4. Arrange: target row ID=" & TARGET_MAT_ID & " with original Fecha=#01/10/2026#"
    logs(4) = "5. Arrange: unrelated guard row ID=" & UNRELATED_MAT_ID & ""
    logs(5) = "6. Arrange: count target row before RegistrarCambioFechaMaterializacion"
    logs(6) = "7. Act: set object Fecha stale, then call RegistrarCambioFechaMaterializacion(#02/20/2026#)"
    logs(7) = "8. Assert: target persisted Fecha equals explicit new date"
    logs(8) = "9. Assert: target cardinality unchanged (1 -> 1)"
    logs(9) = "10. Assert: target unrelated decision/lifecycle fields unchanged"
    logs(10) = "11. Assert/Teardown: unrelated row unchanged; delete deterministic rows"

    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = _
            BuildFail(errMsg, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT CodigoRiesgo FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = _
            BuildFail("Fixture riesgo " & Test_Fixtures.Cache_RiesgoId & " not found", logs)
        GoTo Teardown
    End If
    codigoRiesgo = CStr(Nz(rs.fields("CodigoRiesgo").value, ""))
    rs.Close: Set rs = Nothing
    If codigoRiesgo = "" Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = _
            BuildFail("Fixture riesgo CodigoRiesgo is empty", logs)
        GoTo Teardown
    End If

    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID IN (" & _
        TARGET_MAT_ID & ", " & UNRELATED_MAT_ID & ")", dbFailOnError
    Err.Clear
    On Error GoTo EH

    originalFecha = DateSerial(2026, 1, 10)
    newFecha = DateSerial(2026, 2, 20)
    targetEstado = "Materializado"
    unrelatedFechaBefore = DateSerial(2026, 3, 1)
    unrelatedEstadoBefore = "Materializado"
    unrelatedParaNCBefore = "No"
    unrelatedFechaDecisionBefore = DateSerial(2026, 3, 2)

    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, Estado) " & _
        "VALUES (" & TARGET_MAT_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "#01/10/2026#, 'Sí', '" & targetEstado & "')", dbFailOnError

    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, " & _
        "Estado, ParaNC, FechaDecison) " & _
        "VALUES (" & UNRELATED_MAT_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "#03/01/2026#, 'Sí', '" & unrelatedEstadoBefore & "', 'No', #03/02/2026#)", _
        dbFailOnError

    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & TARGET_MAT_ID, dbOpenSnapshot)
    rowsForTargetBefore = CLng(Nz(rs.fields("Cnt").value, 0))
    rs.Close: Set rs = Nothing
    If rowsForTargetBefore <> 1 Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = _
            BuildFail("Initial cardinality: expected exactly 1 target row, got " & _
                rowsForTargetBefore, logs)
        GoTo Teardown
    End If

    errMsg = ""
    Set mat = Constructor.getRiesgoMaterializado(CStr(TARGET_MAT_ID), errMsg)
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = _
            BuildFail("getRiesgoMaterializado failed: " & errMsg, logs)
        GoTo Teardown
    End If
    If mat Is Nothing Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = _
            BuildFail("getRiesgoMaterializado returned Nothing for ID=" & TARGET_MAT_ID, logs)
        GoTo Teardown
    End If

    mat.Fecha = CStr(originalFecha)
    errMsg = ""
    mat.RegistrarCambioFechaMaterializacion CStr(newFecha), errMsg
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = _
            BuildFail("RegistrarCambioFechaMaterializacion should not fail: " & errMsg, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, " & _
        "Estado, IDNC, ParaNC, FechaDecison FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & TARGET_MAT_ID, dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = _
            BuildFail("Target materialization row not found for ID=" & TARGET_MAT_ID, logs)
        GoTo Teardown
    End If

    If CDate(rs.fields("Fecha").value) <> newFecha Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = _
            BuildFail("Fecha should equal #02/20/2026#, got " & rs.fields("Fecha").value, logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("Fecha").value) = originalFecha Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = _
            BuildFail("Fecha remained at stale object date #01/10/2026#", logs)
        GoTo Teardown
    End If
    If CLng(rs.fields("IDProyecto").value) <> Test_Fixtures.Cache_ProyectoId Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = BuildFail("IDProyecto changed", logs)
        GoTo Teardown
    End If
    If CLng(rs.fields("IDEdicion").value) <> Test_Fixtures.Cache_EdicionId Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = BuildFail("IDEdicion changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("CodigoRiesgo").value, "")), codigoRiesgo, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = BuildFail("CodigoRiesgo changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("EsMaterializacion").value, "")), "Sí", vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = BuildFail("EsMaterializacion changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("Estado").value, "")), targetEstado, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = BuildFail("Estado changed", logs)
        GoTo Teardown
    End If
    If Not IsNull(rs.fields("IDNC").value) Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = BuildFail("IDNC should remain Null", logs)
        GoTo Teardown
    End If
    If Not IsNull(rs.fields("ParaNC").value) Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = BuildFail("ParaNC should remain Null", logs)
        GoTo Teardown
    End If
    If Not IsNull(rs.fields("FechaDecison").value) Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = BuildFail("FechaDecison should remain Null", logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & TARGET_MAT_ID, dbOpenSnapshot)
    rowsForTargetAfter = CLng(Nz(rs.fields("Cnt").value, 0))
    rs.Close: Set rs = Nothing
    If rowsForTargetAfter <> rowsForTargetBefore Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = _
            BuildFail("Cardinality changed: expected " & rowsForTargetBefore & _
                ", got " & rowsForTargetAfter, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT Fecha, Estado, ParaNC, FechaDecison FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & UNRELATED_MAT_ID, dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = _
            BuildFail("Unrelated guard row disappeared", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("Fecha").value) <> unrelatedFechaBefore Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = BuildFail("Unrelated Fecha changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("Estado").value, "")), unrelatedEstadoBefore, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = BuildFail("Unrelated Estado changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("ParaNC").value, "")), unrelatedParaNCBefore, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = BuildFail("Unrelated ParaNC changed", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("FechaDecison").value) <> unrelatedFechaDecisionBefore Then
        Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = BuildFail("Unrelated FechaDecison changed", logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = _
        BuildOk("registrar_cambio_fecha_materializacion_actualiza_fecha", logs)

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set mat = Nothing
    Set cleanupDb = Test_Fixtures.GetTestDb(cleanupErr)
    If Not cleanupDb Is Nothing Then
        cleanupDb.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID IN (" & _
            TARGET_MAT_ID & ", " & UNRELATED_MAT_ID & ")"
        Set cleanupDb = Nothing
    End If
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
    Exit Function

EH:
    Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha = _
        BuildFail("Test_RiesgoMaterializacion_RegistrarCambioFechaMaterializacion_ActualizaFecha: " & _
            Err.description, logs)
    Resume Teardown
End Function

Public Function Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC() As String
    On Error GoTo EH

    Dim logs(0 To 10) As String
    Dim errMsg As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim mat As RiesgoMaterializacion
    Dim codigoRiesgo As String
    Dim rowsForTargetBefore As Long
    Dim rowsForTargetAfter As Long
    Dim targetFechaBefore As Date
    Dim targetEstadoBefore As String
    Dim guardFechaBefore As Date
    Dim guardEstadoBefore As String
    Dim guardParaNCBefore As String
    Dim guardFechaDecisionBefore As Date
    Dim guardIDNCBefore As Long
    Dim cleanupDb As DAO.Database
    Dim cleanupErr As String
    Const TARGET_MAT_ID As Long = 910920
    Const GUARD_MAT_ID As Long = 910922

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: SeedAll fixture graph (project=900501, riesgo=900503)"
    logs(2) = "3. Arrange: resolve fixture CodigoRiesgo without SELECT TOP 1"
    logs(3) = "4. Arrange: target row ID=" & TARGET_MAT_ID & " with non-null IDNC, ParaNC, FechaDecison"
    logs(4) = "5. Arrange: unrelated guard row ID=" & GUARD_MAT_ID & " with its own decision fields"
    logs(5) = "6. Arrange: count target row before RegistrarPorDecidir"
    logs(6) = "7. Act: RegistrarPorDecidir"
    logs(7) = "8. Assert: target IDNC, ParaNC, FechaDecison are Null"
    logs(8) = "9. Assert: target cardinality unchanged (1 -> 1)"
    logs(9) = "10. Assert: unrelated lifecycle fields unchanged"
    logs(10) = "11. Assert/Teardown: guard row unchanged; delete deterministic rows"

    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = _
            BuildFail(errMsg, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT CodigoRiesgo FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = _
            BuildFail("Fixture riesgo " & Test_Fixtures.Cache_RiesgoId & " not found", logs)
        GoTo Teardown
    End If
    codigoRiesgo = CStr(Nz(rs.fields("CodigoRiesgo").value, ""))
    rs.Close: Set rs = Nothing
    If codigoRiesgo = "" Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = _
            BuildFail("Fixture riesgo CodigoRiesgo is empty", logs)
        GoTo Teardown
    End If

    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID IN (" & _
        TARGET_MAT_ID & ", " & GUARD_MAT_ID & ")", dbFailOnError
    Err.Clear
    On Error GoTo EH

    targetFechaBefore = DateSerial(2026, 4, 1)
    targetEstadoBefore = "Materializado"
    guardFechaBefore = DateSerial(2026, 4, 5)
    guardEstadoBefore = "Materializado"
    guardParaNCBefore = "No"
    guardFechaDecisionBefore = DateSerial(2026, 4, 6)
    guardIDNCBefore = 880922

    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, " & _
        "Estado, IDNC, ParaNC, FechaDecison) " & _
        "VALUES (" & TARGET_MAT_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "#04/01/2026#, 'Sí', '" & targetEstadoBefore & "', 880920, 'Sí', #04/02/2026#)", _
        dbFailOnError

    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, " & _
        "Estado, IDNC, ParaNC, FechaDecison) " & _
        "VALUES (" & GUARD_MAT_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "#04/05/2026#, 'Sí', '" & guardEstadoBefore & "', " & _
        guardIDNCBefore & ", '" & guardParaNCBefore & "', #04/06/2026#)", _
        dbFailOnError

    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & TARGET_MAT_ID, dbOpenSnapshot)
    rowsForTargetBefore = CLng(Nz(rs.fields("Cnt").value, 0))
    rs.Close: Set rs = Nothing
    If rowsForTargetBefore <> 1 Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = _
            BuildFail("Initial cardinality: expected exactly 1 target row, got " & _
                rowsForTargetBefore, logs)
        GoTo Teardown
    End If

    errMsg = ""
    Set mat = Constructor.getRiesgoMaterializado(CStr(TARGET_MAT_ID), errMsg)
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = _
            BuildFail("getRiesgoMaterializado failed: " & errMsg, logs)
        GoTo Teardown
    End If
    If mat Is Nothing Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = _
            BuildFail("getRiesgoMaterializado returned Nothing for ID=" & TARGET_MAT_ID, logs)
        GoTo Teardown
    End If

    errMsg = ""
    mat.RegistrarPorDecidir errMsg
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = _
            BuildFail("RegistrarPorDecidir should not fail: " & errMsg, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, " & _
        "Estado, IDNC, ParaNC, FechaDecison FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & TARGET_MAT_ID, dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = _
            BuildFail("Target materialization row not found for ID=" & TARGET_MAT_ID, logs)
        GoTo Teardown
    End If

    If Not IsNull(rs.fields("IDNC").value) Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = BuildFail("IDNC should be Null", logs)
        GoTo Teardown
    End If
    If Not IsNull(rs.fields("ParaNC").value) Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = BuildFail("ParaNC should be Null", logs)
        GoTo Teardown
    End If
    If Not IsNull(rs.fields("FechaDecison").value) Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = BuildFail("FechaDecison should be Null", logs)
        GoTo Teardown
    End If
    If CLng(rs.fields("IDProyecto").value) <> Test_Fixtures.Cache_ProyectoId Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = BuildFail("IDProyecto changed", logs)
        GoTo Teardown
    End If
    If CLng(rs.fields("IDEdicion").value) <> Test_Fixtures.Cache_EdicionId Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = BuildFail("IDEdicion changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("CodigoRiesgo").value, "")), codigoRiesgo, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = BuildFail("CodigoRiesgo changed", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("Fecha").value) <> targetFechaBefore Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = BuildFail("Fecha changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("EsMaterializacion").value, "")), "Sí", vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = BuildFail("EsMaterializacion changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("Estado").value, "")), targetEstadoBefore, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = BuildFail("Estado changed", logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & TARGET_MAT_ID, dbOpenSnapshot)
    rowsForTargetAfter = CLng(Nz(rs.fields("Cnt").value, 0))
    rs.Close: Set rs = Nothing
    If rowsForTargetAfter <> rowsForTargetBefore Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = _
            BuildFail("Cardinality changed: expected " & rowsForTargetBefore & _
                ", got " & rowsForTargetAfter, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT Fecha, Estado, IDNC, ParaNC, FechaDecison " & _
        "FROM TbRiesgosMaterializaciones WHERE ID=" & GUARD_MAT_ID, dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = _
            BuildFail("Unrelated guard row disappeared", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("Fecha").value) <> guardFechaBefore Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = BuildFail("Guard Fecha changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("Estado").value, "")), guardEstadoBefore, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = BuildFail("Guard Estado changed", logs)
        GoTo Teardown
    End If
    If CLng(Nz(rs.fields("IDNC").value, 0)) <> guardIDNCBefore Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = BuildFail("Guard IDNC changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("ParaNC").value, "")), guardParaNCBefore, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = BuildFail("Guard ParaNC changed", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("FechaDecison").value) <> guardFechaDecisionBefore Then
        Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = BuildFail("Guard FechaDecison changed", logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = _
        BuildOk("registrar_por_decidir_limpia_decision_nc", logs)

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set mat = Nothing
    Set cleanupDb = Test_Fixtures.GetTestDb(cleanupErr)
    If Not cleanupDb Is Nothing Then
        cleanupDb.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID IN (" & _
            TARGET_MAT_ID & ", " & GUARD_MAT_ID & ")"
        Set cleanupDb = Nothing
    End If
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
    Exit Function

EH:
    Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC = _
        BuildFail("Test_RiesgoMaterializacion_RegistrarPorDecidir_LimpiaDecisionNC: " & _
            Err.description, logs)
    Resume Teardown
End Function

Public Function Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC() As String
    On Error GoTo EH

    Dim logs(0 To 10) As String
    Dim errMsg As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim mat As RiesgoMaterializacion
    Dim codigoRiesgo As String
    Dim rowsForTargetBefore As Long
    Dim rowsForTargetAfter As Long
    Dim targetFechaBefore As Date
    Dim targetEstadoBefore As String
    Dim guardFechaBefore As Date
    Dim guardEstadoBefore As String
    Dim guardParaNCBefore As String
    Dim guardFechaDecisionBefore As Date
    Dim guardIDNCBefore As Long
    Dim cleanupDb As DAO.Database
    Dim cleanupErr As String
    Const TARGET_MAT_ID As Long = 910930
    Const GUARD_MAT_ID As Long = 910932

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: SeedAll fixture graph (project=900501, riesgo=900503)"
    logs(2) = "3. Arrange: resolve fixture CodigoRiesgo without SELECT TOP 1"
    logs(3) = "4. Arrange: target row ID=" & TARGET_MAT_ID & " with IDNC, ParaNC='Sí', FechaDecison"
    logs(4) = "5. Arrange: unrelated guard row ID=" & GUARD_MAT_ID & " with its own decision fields"
    logs(5) = "6. Arrange: count target row before DesvincularNC"
    logs(6) = "7. Act: DesvincularNC"
    logs(7) = "8. Assert: target IDNC, ParaNC, FechaDecison are Null"
    logs(8) = "9. Assert: target cardinality unchanged (1 -> 1)"
    logs(9) = "10. Assert: unrelated lifecycle fields unchanged"
    logs(10) = "11. Assert/Teardown: guard row unchanged; delete deterministic rows"

    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = _
            BuildFail(errMsg, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT CodigoRiesgo FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = _
            BuildFail("Fixture riesgo " & Test_Fixtures.Cache_RiesgoId & " not found", logs)
        GoTo Teardown
    End If
    codigoRiesgo = CStr(Nz(rs.fields("CodigoRiesgo").value, ""))
    rs.Close: Set rs = Nothing
    If codigoRiesgo = "" Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = _
            BuildFail("Fixture riesgo CodigoRiesgo is empty", logs)
        GoTo Teardown
    End If

    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID IN (" & _
        TARGET_MAT_ID & ", " & GUARD_MAT_ID & ")", dbFailOnError
    Err.Clear
    On Error GoTo EH

    targetFechaBefore = DateSerial(2026, 5, 1)
    targetEstadoBefore = "Materializado"
    guardFechaBefore = DateSerial(2026, 5, 5)
    guardEstadoBefore = "Materializado"
    guardParaNCBefore = "No"
    guardFechaDecisionBefore = DateSerial(2026, 5, 6)
    guardIDNCBefore = 880932

    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, " & _
        "Estado, IDNC, ParaNC, FechaDecison) " & _
        "VALUES (" & TARGET_MAT_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "#05/01/2026#, 'Sí', '" & targetEstadoBefore & "', 880930, 'Sí', #05/02/2026#)", _
        dbFailOnError

    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, " & _
        "Estado, IDNC, ParaNC, FechaDecison) " & _
        "VALUES (" & GUARD_MAT_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "#05/05/2026#, 'Sí', '" & guardEstadoBefore & "', " & _
        guardIDNCBefore & ", '" & guardParaNCBefore & "', #05/06/2026#)", _
        dbFailOnError

    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & TARGET_MAT_ID, dbOpenSnapshot)
    rowsForTargetBefore = CLng(Nz(rs.fields("Cnt").value, 0))
    rs.Close: Set rs = Nothing
    If rowsForTargetBefore <> 1 Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = _
            BuildFail("Initial cardinality: expected exactly 1 target row, got " & _
                rowsForTargetBefore, logs)
        GoTo Teardown
    End If

    errMsg = ""
    Set mat = Constructor.getRiesgoMaterializado(CStr(TARGET_MAT_ID), errMsg)
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = _
            BuildFail("getRiesgoMaterializado failed: " & errMsg, logs)
        GoTo Teardown
    End If
    If mat Is Nothing Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = _
            BuildFail("getRiesgoMaterializado returned Nothing for ID=" & TARGET_MAT_ID, logs)
        GoTo Teardown
    End If

    errMsg = ""
    mat.DesvincularNC errMsg
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = _
            BuildFail("DesvincularNC should not fail: " & errMsg, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, " & _
        "Estado, IDNC, ParaNC, FechaDecison FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & TARGET_MAT_ID, dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = _
            BuildFail("Target materialization row not found for ID=" & TARGET_MAT_ID, logs)
        GoTo Teardown
    End If

    If Not IsNull(rs.fields("IDNC").value) Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = BuildFail("IDNC should be Null", logs)
        GoTo Teardown
    End If
    If Not IsNull(rs.fields("ParaNC").value) Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = BuildFail("ParaNC should be Null", logs)
        GoTo Teardown
    End If
    If Not IsNull(rs.fields("FechaDecison").value) Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = BuildFail("FechaDecison should be Null", logs)
        GoTo Teardown
    End If
    If CLng(rs.fields("IDProyecto").value) <> Test_Fixtures.Cache_ProyectoId Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = BuildFail("IDProyecto changed", logs)
        GoTo Teardown
    End If
    If CLng(rs.fields("IDEdicion").value) <> Test_Fixtures.Cache_EdicionId Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = BuildFail("IDEdicion changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("CodigoRiesgo").value, "")), codigoRiesgo, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = BuildFail("CodigoRiesgo changed", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("Fecha").value) <> targetFechaBefore Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = BuildFail("Fecha changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("EsMaterializacion").value, "")), "Sí", vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = BuildFail("EsMaterializacion changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("Estado").value, "")), targetEstadoBefore, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = BuildFail("Estado changed", logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & TARGET_MAT_ID, dbOpenSnapshot)
    rowsForTargetAfter = CLng(Nz(rs.fields("Cnt").value, 0))
    rs.Close: Set rs = Nothing
    If rowsForTargetAfter <> rowsForTargetBefore Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = _
            BuildFail("Cardinality changed: expected " & rowsForTargetBefore & _
                ", got " & rowsForTargetAfter, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT Fecha, Estado, IDNC, ParaNC, FechaDecison " & _
        "FROM TbRiesgosMaterializaciones WHERE ID=" & GUARD_MAT_ID, dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = _
            BuildFail("Unrelated guard row disappeared", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("Fecha").value) <> guardFechaBefore Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = BuildFail("Guard Fecha changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("Estado").value, "")), guardEstadoBefore, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = BuildFail("Guard Estado changed", logs)
        GoTo Teardown
    End If
    If CLng(Nz(rs.fields("IDNC").value, 0)) <> guardIDNCBefore Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = BuildFail("Guard IDNC changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("ParaNC").value, "")), guardParaNCBefore, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = BuildFail("Guard ParaNC changed", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("FechaDecison").value) <> guardFechaDecisionBefore Then
        Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = BuildFail("Guard FechaDecison changed", logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = _
        BuildOk("desvincular_nc_limpia_decision_nc", logs)

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set mat = Nothing
    Set cleanupDb = Test_Fixtures.GetTestDb(cleanupErr)
    If Not cleanupDb Is Nothing Then
        cleanupDb.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID IN (" & _
            TARGET_MAT_ID & ", " & GUARD_MAT_ID & ")"
        Set cleanupDb = Nothing
    End If
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
    Exit Function

EH:
    Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC = _
        BuildFail("Test_RiesgoMaterializacion_DesvincularNC_LimpiaDecisionNC: " & _
            Err.description, logs)
    Resume Teardown
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

Public Function Test_RiesgoMaterializacion_FechaFuturaBloqueada() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: fecha de materializacion futura"
    logs(1) = "2. Act: ValidarFechaMaterializacionPermitida"
    logs(2) = "3. Assert: devuelve False"
    logs(3) = "4. Assert: informa error funcional"

    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_FechaFutura As String
    Dim m_Resultado As Boolean

    Set m_Riesgo = New riesgo
    m_FechaFutura = Format$(DateAdd("d", 1, Date), "dd/mm/yyyy")

    m_Resultado = m_Riesgo.ValidarFechaMaterializacionPermitida(m_FechaFutura, m_Error)

    If m_Resultado Then
        Test_RiesgoMaterializacion_FechaFuturaBloqueada = BuildFail("La fecha futura debe bloquearse", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "posterior", vbTextCompare) = 0 Then
        Test_RiesgoMaterializacion_FechaFuturaBloqueada = BuildFail("El error debe indicar que la fecha no puede ser posterior a hoy", logs)
        Exit Function
    End If

    Test_RiesgoMaterializacion_FechaFuturaBloqueada = BuildOk("fecha_futura_bloqueada", logs)
    Exit Function
EH:
    Test_RiesgoMaterializacion_FechaFuturaBloqueada = BuildFail("Test_RiesgoMaterializacion_FechaFuturaBloqueada: " & Err.description, logs)
End Function

Public Function Test_RiesgoMaterializacion_FechaHoyPermitida() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: fecha de materializacion igual a hoy"
    logs(1) = "2. Act: ValidarFechaMaterializacionPermitida"
    logs(2) = "3. Assert: devuelve True sin error"

    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_Resultado As Boolean

    Set m_Riesgo = New riesgo

    m_Resultado = m_Riesgo.ValidarFechaMaterializacionPermitida(Format$(Date, "dd/mm/yyyy"), m_Error)

    If Not m_Resultado Then
        Test_RiesgoMaterializacion_FechaHoyPermitida = BuildFail("La fecha de hoy debe permitirse: " & m_Error, logs)
        Exit Function
    End If

    If m_Error <> "" Then
        Test_RiesgoMaterializacion_FechaHoyPermitida = BuildFail("No debe informar error para fecha de hoy: " & m_Error, logs)
        Exit Function
    End If

    Test_RiesgoMaterializacion_FechaHoyPermitida = BuildOk("fecha_hoy_permitida", logs)
    Exit Function
EH:
    Test_RiesgoMaterializacion_FechaHoyPermitida = BuildFail("Test_RiesgoMaterializacion_FechaHoyPermitida: " & Err.description, logs)
End Function

Public Function Test_RiesgoMaterializacion_FechaPasadaPermitida() As String
    On Error GoTo EH

    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: fecha de materializacion anterior a hoy"
    logs(1) = "2. Act: ValidarFechaMaterializacionPermitida"
    logs(2) = "3. Assert: devuelve True sin error"

    Dim m_Riesgo As riesgo
    Dim m_Error As String
    Dim m_FechaPasada As String
    Dim m_Resultado As Boolean

    Set m_Riesgo = New riesgo
    m_FechaPasada = Format$(DateAdd("d", -1, Date), "dd/mm/yyyy")

    m_Resultado = m_Riesgo.ValidarFechaMaterializacionPermitida(m_FechaPasada, m_Error)

    If Not m_Resultado Then
        Test_RiesgoMaterializacion_FechaPasadaPermitida = BuildFail("La fecha pasada debe permitirse: " & m_Error, logs)
        Exit Function
    End If

    If m_Error <> "" Then
        Test_RiesgoMaterializacion_FechaPasadaPermitida = BuildFail("No debe informar error para fecha pasada: " & m_Error, logs)
        Exit Function
    End If

    Test_RiesgoMaterializacion_FechaPasadaPermitida = BuildOk("fecha_pasada_permitida", logs)
    Exit Function
EH:
    Test_RiesgoMaterializacion_FechaPasadaPermitida = BuildFail("Test_RiesgoMaterializacion_FechaPasadaPermitida: " & Err.description, logs)
End Function

Public Function Test_RiesgoMaterializacion_RegistrarFechaFuturaNoPersiste() As String
    On Error GoTo EH

    Dim logs(0 To 5) As String
    logs(0) = "1. Arrange: grafo base y riesgo fixture"
    logs(1) = "2. Arrange: fecha futura"
    logs(2) = "3. Act: MaterializacionRegistrar"
    logs(3) = "4. Assert: devuelve error funcional"
    logs(4) = "5. Assert: no persiste FechaMaterializado"
    logs(5) = "6. Teardown: limpia fixture"

    Dim m_Error As String
    Dim m_DbErr As String
    Dim m_Db As DAO.Database
    Dim m_Riesgo As riesgo
    Dim m_FechaFutura As String
    Dim m_Rs As DAO.Recordset

    ResetGlobals m_Error
    If m_Error <> "" Then
        Test_RiesgoMaterializacion_RegistrarFechaFuturaNoPersiste = _
            BuildFail("ResetGlobals before arrange failed: " & m_Error, logs)
        GoTo Teardown
    End If

    If Not Test_Helper.ForceLocalBackend(m_Error) Then
        Test_RiesgoMaterializacion_RegistrarFechaFuturaNoPersiste = _
            BuildFail("ForceLocalBackend before arrange failed: " & m_Error, logs)
        GoTo Teardown
    End If

    Test_Fixtures.SeedAll
    Set m_Db = Test_Fixtures.GetTestDb(m_DbErr)
    If m_Db Is Nothing Then
        Test_RiesgoMaterializacion_RegistrarFechaFuturaNoPersiste = BuildFail("TESTS BLOCKED: " & m_DbErr, logs)
        GoTo Teardown
    End If

    Set m_Riesgo = Constructor.getRiesgo(p_IDRiesgo:=CStr(Test_Fixtures.Cache_RiesgoId), p_Error:=m_Error)
    If m_Error <> "" Then
        Test_RiesgoMaterializacion_RegistrarFechaFuturaNoPersiste = BuildFail("Constructor.getRiesgo: " & m_Error, logs)
        GoTo Teardown
    End If

    If m_Riesgo Is Nothing Then
        Test_RiesgoMaterializacion_RegistrarFechaFuturaNoPersiste = BuildFail("Constructor.getRiesgo devolvió Nothing", logs)
        GoTo Teardown
    End If

    m_FechaFutura = Format$(DateAdd("d", 1, Date), "dd/mm/yyyy")
    m_Error = ""
    On Error Resume Next
    m_Riesgo.MaterializacionRegistrar m_FechaFutura, m_Error
    Err.Clear
    On Error GoTo EH

    If m_Error = "" Then
        Test_RiesgoMaterializacion_RegistrarFechaFuturaNoPersiste = BuildFail("MaterializacionRegistrar debe informar error para fecha futura", logs)
        GoTo Teardown
    End If

    Set m_Rs = m_Db.OpenRecordset("SELECT FechaMaterializado, Estado FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId)
    If m_Rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarFechaFuturaNoPersiste = BuildFail("No se encontró el riesgo fixture", logs)
        GoTo Teardown
    End If

    If Not IsNull(m_Rs.fields("FechaMaterializado").value) Then
        Test_RiesgoMaterializacion_RegistrarFechaFuturaNoPersiste = BuildFail("FechaMaterializado no debe persistirse con fecha futura", logs)
        GoTo Teardown
    End If

    If StrComp(Nz(m_Rs.fields("Estado").value, ""), "Materializado", vbTextCompare) = 0 Then
        Test_RiesgoMaterializacion_RegistrarFechaFuturaNoPersiste = BuildFail("Estado no debe cambiar a Materializado con fecha futura", logs)
        GoTo Teardown
    End If

    Test_RiesgoMaterializacion_RegistrarFechaFuturaNoPersiste = BuildOk("fecha_futura_no_persiste", logs)

Teardown:
    On Error Resume Next
    If Not m_Rs Is Nothing Then m_Rs.Close
    Set m_Rs = Nothing
    Set m_Riesgo = Nothing
    Set m_Db = Nothing
    Test_Fixtures.TeardownAll
    m_Error = ""
    ResetGlobals m_Error
    On Error GoTo 0
    Exit Function
EH:
    Test_RiesgoMaterializacion_RegistrarFechaFuturaNoPersiste = BuildFail("Test_RiesgoMaterializacion_RegistrarFechaFuturaNoPersiste: " & Err.description, logs)
    Resume Teardown
End Function

' ============================================================
' F.1 — RegistrarAlta happy path (issue #33)
' Persists the materialization row with EsMaterializacion='Sí',
' Estado='Materializado', Fecha=Date, propagates IDPlanContingencia,
' and assigns mat.ID via DameID.
' ============================================================
Public Function Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo() As String
    On Error GoTo EH

    Dim logs(0 To 8) As String
    Dim errMsg As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim mat As RiesgoMaterializacion
    Dim codigoRiesgo As String
    Dim migrationResult As EnumSiNo
    Dim matIdLong As Long
    Dim hasPlan As Boolean
    Dim countBefore As Long   ' cardinalidad skill v2.4 §4.5 — OBLIGATORIO para INSERT
    Dim countAfter As Long
    Dim campo As Variant
    Dim cleanupDb As DAO.Database
    Dim cleanupErr As String

    matIdLong = 0

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: SeedAll fixture graph (project=900501, edicion=900502, riesgo=900503, pc=900505)"
    logs(2) = "3. Arrange: ensure IDPlanContingencia column via MigracionMaterializacionPlanContingencia_AsegurarCampo"
    logs(3) = "4. Arrange: build RiesgoMaterializacion with EsMaterializacion='Sí'"
    logs(4) = "5. Act: RegistrarAlta"
    logs(5) = "6. Assert: errMsg='' and mat.ID is assigned"
    logs(6) = "7. Assert: persisted row matches expected fields (EsMaterializacion, Estado, IDEdicion, CodigoRiesgo, Fecha, IDPlanContingencia)"
    logs(7) = "8. Assert: in-memory derived properties (Estado, EsMaterializacionCalcuado, ParaNCCalculado, Decidido, ColCampos)"
    logs(8) = "9. Teardown: delete inserted row + TeardownAll"

    ' 1. ForceLocalBackend
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail(errMsg, logs)
        Exit Function
    End If

    ' 2. SeedAll
    Test_Fixtures.SeedAll

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        GoTo Teardown
    End If

    ' 3. Ensure IDPlanContingencia column (idempotent migration)
    migrationResult = MigracionMaterializacionPlanContingencia_AsegurarCampo(errMsg)
    If migrationResult <> EnumSiNo.Sí Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("Migration must succeed: " & errMsg, logs)
        GoTo Teardown
    End If
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("Migration must not return errMsg: " & errMsg, logs)
        GoTo Teardown
    End If

    ' 4. Resolve CodigoRiesgo from the fixture riesgo (do not hardcode)
    Set rs = db.OpenRecordset( _
        "SELECT CodigoRiesgo FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("Fixture riesgo " & Test_Fixtures.Cache_RiesgoId & " not found", logs)
        GoTo Teardown
    End If
    codigoRiesgo = CStr(Nz(rs.fields("CodigoRiesgo").value, ""))
    rs.Close: Set rs = Nothing
    If codigoRiesgo = "" Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("Fixture riesgo CodigoRiesgo is empty", logs)
        GoTo Teardown
    End If

    ' 5. Build the materialization object and register
    Set mat = New RiesgoMaterializacion
    mat.IDProyecto = CStr(Test_Fixtures.Cache_ProyectoId)
    mat.IDEdicion = CStr(Test_Fixtures.Cache_EdicionId)
    mat.codigoRiesgo = codigoRiesgo
    mat.EsMaterializacion = "Sí"
    mat.Fecha = CStr(Date)
    mat.ParaNC = ""
    mat.IDNC = ""
    mat.idPlanContingencia = CStr(Test_Fixtures.Cache_PCId)

    ' -- Cardinalidad ANTES (skill v2.4 §4.5 — OBLIGATORIO) -------------
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId & _
        " AND CodigoRiesgo='" & Replace(codigoRiesgo, "'", "''") & "'", _
        dbOpenSnapshot)
    If rs.EOF Then
        countBefore = 0
    Else
        countBefore = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close: Set rs = Nothing
    logs(4) = "5. Act: RegistrarAlta (countBefore=" & countBefore & ")"

    mat.RegistrarAlta p_Error:=errMsg
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("RegistrarAlta should not fail: " & errMsg, logs)
        GoTo Teardown
    End If
    If mat.ID = "" Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("RegistrarAlta should assign mat.ID", logs)
        GoTo Teardown
    End If
    matIdLong = CLng(mat.ID)

    ' -- Cardinalidad DESPUÉS (skill v2.4 §4.5) — countAfter = countBefore + 1 --
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId & _
        " AND CodigoRiesgo='" & Replace(codigoRiesgo, "'", "''") & "'", _
        dbOpenSnapshot)
    If rs.EOF Then
        countAfter = 0
    Else
        countAfter = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close: Set rs = Nothing
    If countAfter <> countBefore + 1 Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("Cardinalidad: expected countAfter=" & (countBefore + 1) & _
                ", got " & countAfter, logs)
        GoTo Teardown
    End If
    logs(5) = "6. Assert: errMsg='' and mat.ID is assigned (countAfter=" & countAfter & ", +1 OK)"

    ' 6. Assert the persisted row
    Set rs = db.OpenRecordset( _
        "SELECT EsMaterializacion, Estado, IDEdicion, CodigoRiesgo, Fecha, IDPlanContingencia " & _
        "FROM TbRiesgosMaterializaciones WHERE ID=" & matIdLong, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("Persisted materialization row not found for ID=" & matIdLong, logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("EsMaterializacion").value, "")), "Sí", vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("EsMaterializacion on row should be 'Sí'", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("Estado").value, "")), "Materializado", vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("Estado on row should be 'Materializado'", logs)
        GoTo Teardown
    End If
    If CLng(Nz(rs.fields("IDEdicion").value, 0)) <> Test_Fixtures.Cache_EdicionId Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("IDEdicion on row should be " & Test_Fixtures.Cache_EdicionId, logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("CodigoRiesgo").value, "")), codigoRiesgo, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("CodigoRiesgo on row should match fixture '" & codigoRiesgo & "'", logs)
        GoTo Teardown
    End If
    If CDate(Nz(rs.fields("Fecha").value, 0)) <> Date Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("Fecha on row should equal today", logs)
        GoTo Teardown
    End If
    If CLng(Nz(rs.fields("IDPlanContingencia").value, 0)) <> Test_Fixtures.Cache_PCId Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("IDPlanContingencia on row should match fixture PC id " & Test_Fixtures.Cache_PCId, logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    ' 7. Assert in-memory derived properties
    If StrComp(CStr(Nz(mat.Estado, "")), "Materializado", vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("mat.Estado should be 'Materializado' in memory, got '" & mat.Estado & "'", logs)
        GoTo Teardown
    End If
    If mat.EsMaterializacionCalcuado <> EnumSiNo.Sí Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("EsMaterializacionCalcuado should be EnumSiNo.Sí", logs)
        GoTo Teardown
    End If
    If mat.ParaNCCalculado <> EnumSiNo.No Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("ParaNCCalculado should be EnumSiNo.No (ParaNC='')", logs)
        GoTo Teardown
    End If
    If mat.Decidido <> EnumSiNo.No Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("Decidido should be EnumSiNo.No (no IDNC, no FechaDecison)", logs)
        GoTo Teardown
    End If

    ' ColCampos should expose IDPlanContingencia
    hasPlan = False
    For Each campo In mat.ColCampos
        If StrComp(CStr(campo), "IDPlanContingencia", vbTextCompare) = 0 Then
            hasPlan = True
            Exit For
        End If
    Next campo
    If Not hasPlan Then
        Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
            BuildFail("ColCampos should expose 'IDPlanContingencia'", logs)
        GoTo Teardown
    End If

    Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
        BuildOk("registrar_alta_persiste_materializado", logs)

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set mat = Nothing
    ' Defensive: remove the row RegistrarAlta inserted, regardless of test outcome
    Set cleanupDb = Test_Fixtures.GetTestDb(cleanupErr)
    If Not cleanupDb Is Nothing Then
        If matIdLong > 0 Then
            cleanupDb.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID=" & matIdLong
        End If
        Set cleanupDb = Nothing
    End If
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
    Exit Function

EH:
    Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo = _
        BuildFail("Test_RiesgoMaterializacion_RegistrarAlta_PersisteMaterializadoYPropagaEstadoARiesgo: " & _
            Err.description, logs)
    Resume Teardown
End Function

' ============================================================
' F.2 — RegistrarAlta sad path: toggle rule rejection (issue #33)
' MotivoNoOKAlta must reject a second 'Sí' materialization when the
' previous state for (IDProyecto, CodigoRiesgo) is already 'Materializado'.
' Uses the validation-only path (MotivoNoOKAlta) instead of RegistrarAlta
' to avoid polluting the parent TbRiesgos.Estado.
' ============================================================
Public Function Test_RiesgoMaterializacion_RegistrarAlta_RechazaToggleRepetido() As String
    On Error GoTo EH

    Dim logs(0 To 8) As String
    Dim errMsg As String
    Dim motivo As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim mat As RiesgoMaterializacion
    Dim codigoRiesgo As String
    Dim migrationResult As EnumSiNo
    Dim rowsForKey As Long
    Dim parentEstado As String
    Dim cleanupDb As DAO.Database
    Dim cleanupErr As String
    Const PREV_ID As Long = 910600

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: SeedAll fixture graph (project=900501, riesgo=900503)"
    logs(2) = "3. Arrange: ensure IDPlanContingencia column via MigracionMaterializacionPlanContingencia_AsegurarCampo"
    logs(3) = "4. Arrange: pre-existing materialization row (ID=" & PREV_ID & ", EsMaterializacion='Sí', Estado='Materializado')"
    logs(4) = "5. Act: MotivoNoOKAlta on a new instance with same EsMaterializacion='Sí'"
    logs(5) = "6. Assert: motivo contains 'No se puede repetir'"
    logs(6) = "7. Assert: no new row was written for this key (rowsForKey=1)"
    logs(7) = "8. Assert: parent TbRiesgos.Estado still 'Detectado'"
    logs(8) = "9. Teardown: delete pre-existing row + TeardownAll"

    ' 1. ForceLocalBackend
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_RiesgoMaterializacion_RegistrarAlta_RechazaToggleRepetido = BuildFail(errMsg, logs)
        Exit Function
    End If

    ' 2. SeedAll
    Test_Fixtures.SeedAll

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoMaterializacion_RegistrarAlta_RechazaToggleRepetido = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        GoTo Teardown
    End If

    ' 3. Ensure IDPlanContingencia column
    migrationResult = MigracionMaterializacionPlanContingencia_AsegurarCampo(errMsg)
    If migrationResult <> EnumSiNo.Sí Then
        Test_RiesgoMaterializacion_RegistrarAlta_RechazaToggleRepetido = _
            BuildFail("Migration must succeed: " & errMsg, logs)
        GoTo Teardown
    End If
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_RegistrarAlta_RechazaToggleRepetido = _
            BuildFail("Migration must not return errMsg: " & errMsg, logs)
        GoTo Teardown
    End If

    ' 4. Resolve CodigoRiesgo from the fixture riesgo (do not hardcode)
    Set rs = db.OpenRecordset( _
        "SELECT CodigoRiesgo FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarAlta_RechazaToggleRepetido = _
            BuildFail("Fixture riesgo " & Test_Fixtures.Cache_RiesgoId & " not found", logs)
        GoTo Teardown
    End If
    codigoRiesgo = CStr(Nz(rs.fields("CodigoRiesgo").value, ""))
    rs.Close: Set rs = Nothing
    If codigoRiesgo = "" Then
        Test_RiesgoMaterializacion_RegistrarAlta_RechazaToggleRepetido = _
            BuildFail("Fixture riesgo CodigoRiesgo is empty", logs)
        GoTo Teardown
    End If

    ' 5. Idempotent pre-existing row
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID=" & PREV_ID, dbFailOnError
    Err.Clear
    On Error GoTo EH

    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, Estado) " & _
        "VALUES (" & PREV_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "Date(), 'Sí', 'Materializado')", dbFailOnError

    ' 6. Act: MotivoNoOKAlta on a new instance with same EsMaterializacion='Sí'.
    '    Using the validation-only path so we do NOT pollute parent TbRiesgos.Estado
    '    (RegistrarAlta would short-circuit before propagation for the "Sí" case
    '    anyway, but the validation path is the canonical contract for the toggle rule).
    Set mat = New RiesgoMaterializacion
    mat.IDProyecto = CStr(Test_Fixtures.Cache_ProyectoId)
    mat.IDEdicion = CStr(Test_Fixtures.Cache_EdicionId)
    mat.codigoRiesgo = codigoRiesgo
    mat.EsMaterializacion = "Sí"
    mat.Fecha = CStr(Date)
    mat.ParaNC = ""
    mat.IDNC = ""

    ' Positional call (no named arg) — safer against binary/disk signature drift.
    ' The class declares 'Optional ByRef p_Error As String' on disk; if the binary
    ' has an older version with a different parameter name, named-arg syntax fails
    ' with "Pocos parámetros. Se esperaba 1." Positional always resolves by index.
    motivo = mat.MotivoNoOKAlta(errMsg)
    ' MotivoNoOKAlta does NOT raise an error on validation failure; it returns
    ' the motivo as the function result. errMsg remains empty unless a DB error
    ' occurred when fetching the previous state.

    If motivo = "" Then
        Test_RiesgoMaterializacion_RegistrarAlta_RechazaToggleRepetido = _
            BuildFail("MotivoNoOKAlta should return a motivo for toggle rule, got empty", logs)
        GoTo Teardown
    End If
    If InStr(1, motivo, "No se puede repetir", vbTextCompare) = 0 Then
        Test_RiesgoMaterializacion_RegistrarAlta_RechazaToggleRepetido = _
            BuildFail("Motivo should contain 'No se puede repetir', got: " & motivo, logs)
        GoTo Teardown
    End If
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_RegistrarAlta_RechazaToggleRepetido = _
            BuildFail("MotivoNoOKAlta should not raise errMsg on validation failure, got: " & errMsg, logs)
        GoTo Teardown
    End If

    ' 7. Assert no new row was written for this key
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId & _
        " AND CodigoRiesgo='" & Replace(codigoRiesgo, "'", "''") & "'", _
        dbOpenSnapshot)
    If rs.EOF Then
        rowsForKey = 0
    Else
        rowsForKey = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close: Set rs = Nothing
    If rowsForKey <> 1 Then
        Test_RiesgoMaterializacion_RegistrarAlta_RechazaToggleRepetido = _
            BuildFail("Expected exactly 1 row for (project=" & Test_Fixtures.Cache_ProyectoId & _
                ", codigo='" & codigoRiesgo & "'), got " & rowsForKey, logs)
        GoTo Teardown
    End If

    ' 8. Assert parent TbRiesgos.Estado is still 'Detectado' (no RegistrarAlta call was made)
    Set rs = db.OpenRecordset( _
        "SELECT Estado FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarAlta_RechazaToggleRepetido = _
            BuildFail("Fixture riesgo " & Test_Fixtures.Cache_RiesgoId & " missing", logs)
        GoTo Teardown
    End If
    parentEstado = CStr(Nz(rs.fields("Estado").value, ""))
    rs.Close: Set rs = Nothing
    If StrComp(parentEstado, "Detectado", vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarAlta_RechazaToggleRepetido = _
            BuildFail("Parent TbRiesgos.Estado should remain 'Detectado', got '" & parentEstado & "'", logs)
        GoTo Teardown
    End If

    Test_RiesgoMaterializacion_RegistrarAlta_RechazaToggleRepetido = _
        BuildOk("toggle_rule_rejected", logs)

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set mat = Nothing
    ' Defensive: delete pre-existing row regardless of test outcome
    Set cleanupDb = Test_Fixtures.GetTestDb(cleanupErr)
    If Not cleanupDb Is Nothing Then
        cleanupDb.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID=" & PREV_ID
        Set cleanupDb = Nothing
    End If
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
    Exit Function

EH:
    Test_RiesgoMaterializacion_RegistrarAlta_RechazaToggleRepetido = _
        BuildFail("Test_RiesgoMaterializacion_RegistrarAlta_RechazaToggleRepetido: " & _
            Err.description, logs)
    Resume Teardown
End Function

' ============================================================
' F.3 — RegistrarParaNC happy path (issue #33)
' Persists the triple (IDNC, ParaNC='Sí', FechaDecison) on a pre-existing
' 'Sí' materialization row when the referenced NC exists.
'
' Notes:
' - Constructor.getNC queries `TbNoConformidades` (NOT `TbRiesgosNC`).
'   `TbNoConformidades` is a LINKED table in the staging backend pointing
'   to `C:\00repos\datos\NoConformidades_Datos.accdb`, so the NC fixture is
'   wrapped in a DAO transaction and rolled back. This keeps the behavior
'   fixture-first without leaving committed rows in the shared linked backend.
' - `getNC` uses `Nz(...,"")` to read all 25 NC fields, so a minimal INSERT
'   with only the PK (`IDNoConformidad`) is sufficient for `getNC` to find
'   the row and assign empty strings to the rest.
' - `RegistrarParaNC` is EDIT-only: it does NOT insert a new row. Cardinalidad
'   must remain constant (1 ? 1) across the call.
' - Positional call (no named args) for binary/disk signature resilience.
' ============================================================
Public Function Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison() As String
    On Error GoTo EH

    Dim logs(0 To 9) As String
    Dim errMsg As String
    Dim fechaDecision As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim mat As RiesgoMaterializacion
    Dim codigoRiesgo As String
    Dim migrationResult As EnumSiNo
    Dim matIdLong As Long
    Dim ncIdLong As Long
    Dim rowsForKeyBefore As Long
    Dim rowsForKeyAfter As Long
    Dim cleanupDb As DAO.Database
    Dim cleanupErr As String
    Dim ws As DAO.Workspace
    Dim transStarted As Boolean
    Dim existingNcCount As Long
    Const PREV_MAT_ID As Long = 910700
    Const FIX_NC_ID As Long = 910701

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: SeedAll fixture graph (project=900501, riesgo=900503)"
    logs(2) = "3. Arrange: ensure IDPlanContingencia column via MigracionMaterializacionPlanContingencia_AsegurarCampo"
    logs(3) = "4. Arrange: pre-existing 'Sí' materialization row at ID=" & PREV_MAT_ID & " (no NC triple yet)"
    logs(4) = "5. Arrange: pre-existing NC row at ID=" & FIX_NC_ID & " in TbNoConformidades (Constructor.getNC seam)"
    logs(5) = "6. Act: RegistrarParaNC(ncId, fechaDecision)"
    logs(6) = "7. Assert: errMsg='' and row triple updated (IDNC=" & FIX_NC_ID & ", ParaNC='Sí', FechaDecison=Date)"
    logs(7) = "8. Assert: in-memory mat triple matches (IDNC=" & FIX_NC_ID & ", ParaNC='Sí', Decidido=Sí)"
    logs(8) = "9. Assert: cardinalidad rowsForKey unchanged (RegistrarParaNC is EDIT-only)"
    logs(9) = "10. Teardown: rollback transaction + TeardownAll"

    ' 1. ForceLocalBackend
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail(errMsg, logs)
        Exit Function
    End If

    ' 2. SeedAll
    Test_Fixtures.SeedAll

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        GoTo Teardown
    End If

    ' 3. Ensure IDPlanContingencia column (idempotent)
    migrationResult = MigracionMaterializacionPlanContingencia_AsegurarCampo(errMsg)
    If migrationResult <> EnumSiNo.Sí Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Migration must succeed: " & errMsg, logs)
        GoTo Teardown
    End If
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Migration must not return errMsg: " & errMsg, logs)
        GoTo Teardown
    End If

    ' 4. Resolve CodigoRiesgo from the fixture riesgo (do not hardcode)
    Set rs = db.OpenRecordset( _
        "SELECT CodigoRiesgo FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Fixture riesgo " & Test_Fixtures.Cache_RiesgoId & " not found", logs)
        GoTo Teardown
    End If
    codigoRiesgo = CStr(Nz(rs.fields("CodigoRiesgo").value, ""))
    rs.Close: Set rs = Nothing
    If codigoRiesgo = "" Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Fixture riesgo CodigoRiesgo is empty", logs)
        GoTo Teardown
    End If

    ' 5. Pre-check shared NC sentinel without mutating it. If the sentinel ever
    '    exists, block instead of deleting a row from the linked shared backend.
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbNoConformidades " & _
        "WHERE IDNoConformidad=" & FIX_NC_ID, _
        dbOpenSnapshot)
    If rs.EOF Then
        existingNcCount = 0
    Else
        existingNcCount = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close: Set rs = Nothing
    If existingNcCount <> 0 Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("TESTS BLOCKED: sentinel NC " & FIX_NC_ID & _
                " exists in linked TbNoConformidades; refusing to delete shared data", logs)
        GoTo Teardown
    End If

    ' 6. Start transaction before touching the linked NC table. Assertions run
    '    inside the transaction; teardown rolls it back even on failures.
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    transStarted = True

    ' 7. Pre-existing 'Sí' materialization row (idempotent INSERT)
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID=" & PREV_MAT_ID, dbFailOnError
    Err.Clear
    On Error GoTo EH

    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, Estado) " & _
        "VALUES (" & PREV_MAT_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "Date(), 'Sí', 'Materializado')", dbFailOnError

    ' 8. Pre-existing NC row in TbNoConformidades (Constructor.getNC queries this table).
    '    The pre-check above guarantees this INSERT does not overwrite/delete
    '    existing shared data; rollback removes it at teardown.
    db.Execute "INSERT INTO TbNoConformidades (IDNoConformidad, CodigoNoConformidad, EXPEDIENTE, Juridica, TIPO, ESTADO, FECHAAPERTURA) " & _
        "VALUES (" & FIX_NC_ID & ", 'TEST_F3_" & FIX_NC_ID & "', " & _
        Test_Fixtures.Cache_ExpedienteId & ", 'TdE', 'General', 'Abierta', Date())", dbFailOnError
    ncIdLong = FIX_NC_ID

    ' 9. Cardinalidad ANTES — should be 1 (the pre-existing mat row)
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId & _
        " AND CodigoRiesgo='" & Replace(codigoRiesgo, "'", "''") & "'", _
        dbOpenSnapshot)
    If rs.EOF Then
        rowsForKeyBefore = 0
    Else
        rowsForKeyBefore = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close: Set rs = Nothing
    If rowsForKeyBefore <> 1 Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Cardinalidad inicial: expected exactly 1 materialization row for fixture key, got " & _
                rowsForKeyBefore, logs)
        GoTo Teardown
    End If

    ' 10. Load the pre-existing materialization via Constructor.getRiesgoMaterializado
    errMsg = ""
    Set mat = Constructor.getRiesgoMaterializado(CStr(PREV_MAT_ID), errMsg)
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("getRiesgoMaterializado failed: " & errMsg, logs)
        GoTo Teardown
    End If
    If mat Is Nothing Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("getRiesgoMaterializado returned Nothing for ID=" & PREV_MAT_ID, logs)
        GoTo Teardown
    End If
    matIdLong = PREV_MAT_ID

    ' 11. Act: RegistrarParaNC (positional: p_IDNC, p_FechaDecision, p_Error)
    '    Using Date() (with parens) in inline SQL via fechaDecision to avoid the
    '    "Pocos parámetros" trap (bare Date is parsed as a parameter placeholder).
    fechaDecision = CStr(Date)
    errMsg = ""
    mat.RegistrarParaNC CStr(ncIdLong), fechaDecision, errMsg
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("RegistrarParaNC should not fail: " & errMsg, logs)
        GoTo Teardown
    End If

    ' 12. Assert row triple on TbRiesgosMaterializaciones
    Set rs = db.OpenRecordset( _
        "SELECT IDNC, ParaNC, FechaDecison FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & matIdLong, dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Materialization row not found for ID=" & matIdLong, logs)
        GoTo Teardown
    End If

    If IsNull(rs.fields("IDNC").value) Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("IDNC on row should be " & ncIdLong & ", got Null", logs)
        GoTo Teardown
    End If
    If CLng(rs.fields("IDNC").value) <> ncIdLong Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("IDNC on row should be " & ncIdLong & ", got " & rs.fields("IDNC").value, logs)
        GoTo Teardown
    End If

    If StrComp(CStr(Nz(rs.fields("ParaNC").value, "")), "Sí", vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("ParaNC on row should be 'Sí', got '" & Nz(rs.fields("ParaNC").value, "") & "'", logs)
        GoTo Teardown
    End If

    If IsNull(rs.fields("FechaDecison").value) Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("FechaDecison on row should be a valid Date, got Null", logs)
        GoTo Teardown
    End If
    If Not IsDate(rs.fields("FechaDecison").value) Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("FechaDecison on row should be a valid Date, got '" & rs.fields("FechaDecison").value & "'", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("FechaDecison").value) <> Date Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("FechaDecison on row should equal today, got " & rs.fields("FechaDecison").value, logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    ' 13. Assert in-memory mat triple
    If mat.IDNC <> CStr(ncIdLong) Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("mat.IDNC in memory should be " & ncIdLong & ", got '" & mat.IDNC & "'", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(mat.ParaNC, "")), "Sí", vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("mat.ParaNC in memory should be 'Sí', got '" & mat.ParaNC & "'", logs)
        GoTo Teardown
    End If
    If Not IsDate(mat.FechaDecison) Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("mat.FechaDecison in memory should be a Date, got '" & mat.FechaDecison & "'", logs)
        GoTo Teardown
    End If
    If mat.Decidido <> EnumSiNo.Sí Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("mat.Decidido should be EnumSiNo.Sí (IDNC, ParaNC set, FechaDecison valid Date)", logs)
        GoTo Teardown
    End If

    ' 14. Assert cardinalidad (RegistrarParaNC is EDIT-only, no new row)
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId & _
        " AND CodigoRiesgo='" & Replace(codigoRiesgo, "'", "''") & "'", _
        dbOpenSnapshot)
    If rs.EOF Then
        rowsForKeyAfter = 0
    Else
        rowsForKeyAfter = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close: Set rs = Nothing
    If rowsForKeyAfter <> 1 Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Cardinalidad final: expected exactly 1 materialization row for fixture key, got " & _
                rowsForKeyAfter, logs)
        GoTo Teardown
    End If
    If rowsForKeyAfter <> rowsForKeyBefore Then
        Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Cardinalidad: expected rowsForKey unchanged (" & rowsForKeyBefore & _
                "), got " & rowsForKeyAfter & " (RegistrarParaNC must be EDIT-only)", logs)
        GoTo Teardown
    End If

    Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
        BuildOk("registrar_para_nc_persiste_triple", logs)

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set mat = Nothing
    If transStarted Then
        ws.Rollback
        transStarted = False
    Else
        ' Defensive fallback only if the transaction was not opened. Never touch
        ' TbNoConformidades here because it is a linked shared backend table.
        Set cleanupDb = Test_Fixtures.GetTestDb(cleanupErr)
        If Not cleanupDb Is Nothing Then
            cleanupDb.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID=" & PREV_MAT_ID
            Set cleanupDb = Nothing
        End If
    End If
    Set ws = Nothing
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
    Exit Function

EH:
    Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison = _
        BuildFail("Test_RiesgoMaterializacion_RegistrarParaNC_PersisteTripleIDNCParaNCFechaDecison: " & _
            Err.description, logs)
    Resume Teardown
End Function

' ============================================================
' F.4 — RegistrarParaNC sad path: NC doesn't exist (issue #33)
' RegistrarParaNC must set p_Error and NOT modify the row when the
' referenced NC ID is not present in TbNoConformidades.
' ============================================================
Public Function Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente() As String
    On Error GoTo EH

    Dim logs(0 To 7) As String
    Dim errMsg As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim mat As RiesgoMaterializacion
    Dim codigoRiesgo As String
    Dim migrationResult As EnumSiNo
    Dim matIdLong As Long
    Dim ncIdInexistente As Long
    Dim rowsForKeyBefore As Long
    Dim rowsForKeyAfter As Long
    Dim idncInRow As Variant
    Dim paraNcInRow As Variant
    Dim fechaDecisonInRow As Variant
    Dim cleanupDb As DAO.Database
    Dim cleanupErr As String
    Dim missingNcCount As Long
    Const PREV_MAT_ID As Long = 910800
    Const NC_INEXISTENTE As Long = 999999

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: SeedAll fixture graph (project=900501, riesgo=900503)"
    logs(2) = "3. Arrange: ensure IDPlanContingencia column via MigracionMaterializacionPlanContingencia_AsegurarCampo"
    logs(3) = "4. Arrange: pre-existing 'Sí' materialization row at ID=" & PREV_MAT_ID & " (NO NC row inserted)"
    logs(4) = "5. Act: RegistrarParaNC(" & NC_INEXISTENTE & ", '')"
    logs(5) = "6. Assert: errMsg is non-empty (Constructor.getNC returns Nothing for missing ID)"
    logs(6) = "7. Assert: row triple unchanged (IDNC=Null, ParaNC=Null, FechaDecison=Null)"
    logs(7) = "8. Teardown: delete pre-existing materialization row + TeardownAll"

    ' 1. ForceLocalBackend
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = BuildFail(errMsg, logs)
        Exit Function
    End If

    ' 2. SeedAll
    Test_Fixtures.SeedAll

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        GoTo Teardown
    End If

    ' 3. Ensure IDPlanContingencia column
    migrationResult = MigracionMaterializacionPlanContingencia_AsegurarCampo(errMsg)
    If migrationResult <> EnumSiNo.Sí Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
            BuildFail("Migration must succeed: " & errMsg, logs)
        GoTo Teardown
    End If
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
            BuildFail("Migration must not return errMsg: " & errMsg, logs)
        GoTo Teardown
    End If

    ' 4. Resolve CodigoRiesgo
    Set rs = db.OpenRecordset( _
        "SELECT CodigoRiesgo FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
            BuildFail("Fixture riesgo " & Test_Fixtures.Cache_RiesgoId & " not found", logs)
        GoTo Teardown
    End If
    codigoRiesgo = CStr(Nz(rs.fields("CodigoRiesgo").value, ""))
    rs.Close: Set rs = Nothing
    If codigoRiesgo = "" Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
            BuildFail("Fixture riesgo CodigoRiesgo is empty", logs)
        GoTo Teardown
    End If

    ' 5. Pre-existing 'Sí' materialization row, NO NC row inserted.
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID=" & PREV_MAT_ID, dbFailOnError
    Err.Clear
    On Error GoTo EH

    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, Estado) " & _
        "VALUES (" & PREV_MAT_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "Date(), 'Sí', 'Materializado')", dbFailOnError

    ' Defensive: ensure the NC really doesn't exist in the linked table without
    ' deleting from the shared NC backend. If the sentinel ID ever exists, the
    ' fixture is unsafe and this test must block instead of mutating real data.
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbNoConformidades " & _
        "WHERE IDNoConformidad=" & NC_INEXISTENTE, _
        dbOpenSnapshot)
    If rs.EOF Then
        missingNcCount = 0
    Else
        missingNcCount = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close: Set rs = Nothing
    If missingNcCount <> 0 Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
            BuildFail("TESTS BLOCKED: sentinel NC " & NC_INEXISTENTE & _
                " exists in linked TbNoConformidades; refusing to delete shared data", logs)
        GoTo Teardown
    End If

    ' 6. Cardinalidad ANTES
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId & _
        " AND CodigoRiesgo='" & Replace(codigoRiesgo, "'", "''") & "'", _
        dbOpenSnapshot)
    If rs.EOF Then
        rowsForKeyBefore = 0
    Else
        rowsForKeyBefore = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close: Set rs = Nothing
    If rowsForKeyBefore <> 1 Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
            BuildFail("Cardinalidad inicial: expected exactly 1 materialization row for fixture key, got " & _
                rowsForKeyBefore, logs)
        GoTo Teardown
    End If

    ' 7. Load materialization
    errMsg = ""
    Set mat = Constructor.getRiesgoMaterializado(CStr(PREV_MAT_ID), errMsg)
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
            BuildFail("getRiesgoMaterializado failed: " & errMsg, logs)
        GoTo Teardown
    End If
    If mat Is Nothing Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
            BuildFail("getRiesgoMaterializado returned Nothing for ID=" & PREV_MAT_ID, logs)
        GoTo Teardown
    End If
    matIdLong = PREV_MAT_ID

    ' 8. Act: RegistrarParaNC with non-existent NC ID
    '    Constructor.getNC will return Nothing (no row found for IDNoConformidad=999999)
    '    RegistrarParaNC then sets p_Error = "No hay ninguna No Conformidad registrada
    '    con el ID 999999" and raises 1000. The local errores: block preserves p_Error
    '    (since Err.Number = 1000). NO write to TbRiesgosMaterializaciones happens.
    ncIdInexistente = NC_INEXISTENTE
    errMsg = ""
    mat.RegistrarParaNC CStr(ncIdInexistente), "", errMsg
    ' Sad path: errMsg MUST be non-empty
    If errMsg = "" Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
            BuildFail("RegistrarParaNC should set errMsg when NC=" & ncIdInexistente & _
                " does not exist, got empty errMsg", logs)
        GoTo Teardown
    End If

    ' 9. Assert row triple is UNCHANGED (IDNC=Null, ParaNC=Null, FechaDecison=Null)
    Set rs = db.OpenRecordset( _
        "SELECT IDNC, ParaNC, FechaDecison FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & matIdLong, dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
            BuildFail("Materialization row not found for ID=" & matIdLong, logs)
        GoTo Teardown
    End If

    idncInRow = rs.fields("IDNC").value
    paraNcInRow = rs.fields("ParaNC").value
    fechaDecisonInRow = rs.fields("FechaDecison").value
    rs.Close: Set rs = Nothing

    If Not IsNull(idncInRow) Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
            BuildFail("IDNC on row should remain Null after rejected RegistrarParaNC, got " & idncInRow, logs)
        GoTo Teardown
    End If
    If Not IsNull(paraNcInRow) Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
            BuildFail("ParaNC on row should remain Null after rejected RegistrarParaNC, got '" & paraNcInRow & "'", logs)
        GoTo Teardown
    End If
    If Not IsNull(fechaDecisonInRow) Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
            BuildFail("FechaDecison on row should remain Null after rejected RegistrarParaNC, got '" & fechaDecisonInRow & "'", logs)
        GoTo Teardown
    End If

    ' 10. Assert cardinalidad (no new row, no edit to existing)
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId & _
        " AND CodigoRiesgo='" & Replace(codigoRiesgo, "'", "''") & "'", _
        dbOpenSnapshot)
    If rs.EOF Then
        rowsForKeyAfter = 0
    Else
        rowsForKeyAfter = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close: Set rs = Nothing
    If rowsForKeyAfter <> 1 Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
            BuildFail("Cardinalidad final: expected exactly 1 materialization row for fixture key, got " & _
                rowsForKeyAfter, logs)
        GoTo Teardown
    End If
    If rowsForKeyAfter <> rowsForKeyBefore Then
        Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
            BuildFail("Cardinalidad: expected rowsForKey unchanged (" & rowsForKeyBefore & _
                "), got " & rowsForKeyAfter, logs)
        GoTo Teardown
    End If

    Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
        BuildOk("registrar_para_nc_rechaza_nc_inexistente", logs)

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set mat = Nothing
    Set cleanupDb = Test_Fixtures.GetTestDb(cleanupErr)
    If Not cleanupDb Is Nothing Then
        cleanupDb.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID=" & PREV_MAT_ID
        Set cleanupDb = Nothing
    End If
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
    Exit Function

EH:
    Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente = _
        BuildFail("Test_RiesgoMaterializacion_RegistrarParaNC_RechazaNCInexistente: " & _
            Err.description, logs)
    Resume Teardown
End Function

' ============================================================
' F.5 — VincularNC happy path (issue #34, PR 5)
' Inverse of F.4/DesvincularNC_LimpiaDecisionNC: pre-Act NC
' triple is Null; VincularNC must persist IDNC, ParaNC='Sí',
' FechaDecison while preserving unrelated fields and the
' guard row. Mirrors F.3 transaction discipline because
' `TbNoConformidades` is a linked shared backend.
' ============================================================
Public Function Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison() As String
    On Error GoTo EH

    Dim logs(0 To 9) As String
    Dim errMsg As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim mat As RiesgoMaterializacion
    Dim nc As nc
    Dim codigoRiesgo As String
    Dim migrationResult As EnumSiNo
    Dim matIdLong As Long
    Dim ncIdLong As Long
    Dim rowsForTargetBefore As Long
    Dim rowsForTargetAfter As Long
    Dim targetFechaBefore As Date
    Dim targetEstadoBefore As String
    Dim guardFechaBefore As Date
    Dim guardEstadoBefore As String
    Dim guardParaNCBefore As String
    Dim guardFechaDecisionBefore As Date
    Dim guardIDNCBefore As Long
    Dim cleanupDb As DAO.Database
    Dim cleanupErr As String
    Dim ws As DAO.Workspace
    Dim transStarted As Boolean
    Dim existingNcCount As Long
    Const TARGET_MAT_ID As Long = 910940
    Const FIX_NC_ID As Long = 910941
    Const GUARD_MAT_ID As Long = 910942

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: SeedAll fixture graph (project=900501, riesgo=900503)"
    logs(2) = "3. Arrange: ensure IDPlanContingencia column via MigracionMaterializacionPlanContingencia_AsegurarCampo"
    logs(3) = "4. Arrange: resolve fixture CodigoRiesgo without SELECT TOP 1"
    logs(4) = "5. Arrange: sentinel pre-check on linked NC table (no deletion of shared data)"
    logs(5) = "6. Arrange: BeginTrans + INSERT target mat row (NC triple Null) + INSERT guard mat row + INSERT NC row in TbNoConformidades"
    logs(6) = "7. Act: VincularNC (inverse of DesvincularNC)"
    logs(7) = "8. Assert: target row triple updated (IDNC=" & FIX_NC_ID & ", ParaNC='Sí', FechaDecison=Date)"
    logs(8) = "9. Assert: in-memory mat triple mirrors; unrelated lifecycle fields + guard row preserved; cardinalidad 1 -> 1"
    logs(9) = "10. Teardown: rollback transaction + defensive delete + TeardownAll"

    ' 1. ForceLocalBackend
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail(errMsg, logs)
        Exit Function
    End If

    ' 2. SeedAll
    Test_Fixtures.SeedAll

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        GoTo Teardown
    End If

    ' 3. Ensure IDPlanContingencia column (idempotent safety net)
    migrationResult = MigracionMaterializacionPlanContingencia_AsegurarCampo(errMsg)
    If migrationResult <> EnumSiNo.Sí Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Migration must succeed: " & errMsg, logs)
        GoTo Teardown
    End If
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Migration must not return errMsg: " & errMsg, logs)
        GoTo Teardown
    End If

    ' 4. Resolve CodigoRiesgo from the fixture riesgo (do not hardcode)
    Set rs = db.OpenRecordset( _
        "SELECT CodigoRiesgo FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Fixture riesgo " & Test_Fixtures.Cache_RiesgoId & " not found", logs)
        GoTo Teardown
    End If
    codigoRiesgo = CStr(Nz(rs.fields("CodigoRiesgo").value, ""))
    rs.Close: Set rs = Nothing
    If codigoRiesgo = "" Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Fixture riesgo CodigoRiesgo is empty", logs)
        GoTo Teardown
    End If

    ' 5. Pre-check shared NC sentinel without mutating it. If the sentinel ever
    '    exists, block instead of deleting a row from the linked shared backend.
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbNoConformidades " & _
        "WHERE IDNoConformidad=" & FIX_NC_ID, _
        dbOpenSnapshot)
    If rs.EOF Then
        existingNcCount = 0
    Else
        existingNcCount = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close: Set rs = Nothing
    If existingNcCount <> 0 Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("TESTS BLOCKED: sentinel NC " & FIX_NC_ID & _
                " exists in linked TbNoConformidades; refusing to delete shared data", logs)
        GoTo Teardown
    End If

    ' 6. Start transaction before touching the linked NC table. Assertions run
    '    inside the transaction; teardown rolls it back even on failures.
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    transStarted = True

    ' 6a. Defensive pre-clean of target/guard materialization rows.
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID IN (" & _
        TARGET_MAT_ID & ", " & GUARD_MAT_ID & ")", dbFailOnError
    Err.Clear
    On Error GoTo EH

    ' 6b. Target row: pre-existing 'Sí' materialization with ALL three NC
    '     decision fields NULL (this is the inverse state of the
    '     DesvincularNC test's arrange block).
    targetFechaBefore = DateSerial(2026, 5, 10)
    targetEstadoBefore = "Materializado"
    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, Estado) " & _
        "VALUES (" & TARGET_MAT_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "#05/10/2026#, 'Sí', '" & targetEstadoBefore & "')", dbFailOnError

    ' 6c. Guard row: pre-existing 'Sí' materialization with its own decision
    '     fields, which VincularNC on the target row MUST NOT touch.
    guardFechaBefore = DateSerial(2026, 5, 15)
    guardEstadoBefore = "Materializado"
    guardParaNCBefore = "No"
    guardFechaDecisionBefore = DateSerial(2026, 5, 16)
    guardIDNCBefore = 880942
    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, " & _
        "Estado, IDNC, ParaNC, FechaDecison) " & _
        "VALUES (" & GUARD_MAT_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "#05/15/2026#, 'Sí', '" & guardEstadoBefore & "', " & _
        guardIDNCBefore & ", '" & guardParaNCBefore & "', #05/16/2026#)", _
        dbFailOnError

    ' 6d. NC row in TbNoConformidades (Constructor.getNC queries this table).
    '     The pre-check above guarantees this INSERT does not overwrite/delete
    '     existing shared data; rollback removes it at teardown.
    db.Execute "INSERT INTO TbNoConformidades (IDNoConformidad, CodigoNoConformidad, EXPEDIENTE, Juridica, TIPO, ESTADO, FECHAAPERTURA) " & _
        "VALUES (" & FIX_NC_ID & ", 'TEST_PR5_" & FIX_NC_ID & "', " & _
        Test_Fixtures.Cache_ExpedienteId & ", 'TdE', 'General', 'Abierta', Date())", dbFailOnError
    ncIdLong = FIX_NC_ID

    ' 7. Cardinalidad ANTES — should be exactly 1 target row
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & TARGET_MAT_ID, dbOpenSnapshot)
    If rs.EOF Then
        rowsForTargetBefore = 0
    Else
        rowsForTargetBefore = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close: Set rs = Nothing
    If rowsForTargetBefore <> 1 Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Cardinalidad inicial: expected exactly 1 target row, got " & _
                rowsForTargetBefore, logs)
        GoTo Teardown
    End If

    ' 8. Load the pre-existing materialization via Constructor.getRiesgoMaterializado.
    '    After load, in-memory triple must be Null (pre-Act state for VincularNC).
    errMsg = ""
    Set mat = Constructor.getRiesgoMaterializado(CStr(TARGET_MAT_ID), errMsg)
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("getRiesgoMaterializado failed: " & errMsg, logs)
        GoTo Teardown
    End If
    If mat Is Nothing Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("getRiesgoMaterializado returned Nothing for ID=" & TARGET_MAT_ID, logs)
        GoTo Teardown
    End If
    matIdLong = TARGET_MAT_ID

    ' 9. Load NC via Constructor.getNC. VincularNC needs the live NC instance.
    errMsg = ""
    Set nc = Constructor.getNC(CStr(FIX_NC_ID), errMsg)
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("getNC failed: " & errMsg, logs)
        GoTo Teardown
    End If
    If nc Is Nothing Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("getNC returned Nothing for ID=" & FIX_NC_ID, logs)
        GoTo Teardown
    End If

    ' 10. Act: VincularNC (inverse of DesvincularNC). Positional call.
    errMsg = ""
    mat.VincularNC nc, errMsg
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("VincularNC should not fail: " & errMsg, logs)
        GoTo Teardown
    End If

    ' 11. Assert target row triple in TbRiesgosMaterializaciones.
    Set rs = db.OpenRecordset( _
        "SELECT IDNC, ParaNC, FechaDecison FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & matIdLong, dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Materialization row not found for ID=" & matIdLong, logs)
        GoTo Teardown
    End If

    If IsNull(rs.fields("IDNC").value) Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("IDNC on row should be " & ncIdLong & ", got Null", logs)
        GoTo Teardown
    End If
    If CLng(rs.fields("IDNC").value) <> ncIdLong Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("IDNC on row should be " & ncIdLong & ", got " & rs.fields("IDNC").value, logs)
        GoTo Teardown
    End If

    If StrComp(CStr(Nz(rs.fields("ParaNC").value, "")), "Sí", vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("ParaNC on row should be 'Sí', got '" & Nz(rs.fields("ParaNC").value, "") & "'", logs)
        GoTo Teardown
    End If

    If IsNull(rs.fields("FechaDecison").value) Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("FechaDecison on row should be a valid Date, got Null", logs)
        GoTo Teardown
    End If
    If Not IsDate(rs.fields("FechaDecison").value) Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("FechaDecison on row should be a valid Date, got '" & rs.fields("FechaDecison").value & "'", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("FechaDecison").value) <> Date Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("FechaDecison on row should equal today, got " & rs.fields("FechaDecison").value, logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    ' 12. Assert in-memory mat triple mirrors the persisted triple.
    If mat.IDNC <> CStr(ncIdLong) Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("mat.IDNC in memory should be " & ncIdLong & ", got '" & mat.IDNC & "'", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(mat.ParaNC, "")), "Sí", vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("mat.ParaNC in memory should be 'Sí', got '" & mat.ParaNC & "'", logs)
        GoTo Teardown
    End If
    If Not IsDate(mat.FechaDecison) Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("mat.FechaDecison in memory should be a Date, got '" & mat.FechaDecison & "'", logs)
        GoTo Teardown
    End If

    ' 13. Assert unrelated lifecycle fields on target row are preserved.
    Set rs = db.OpenRecordset( _
        "SELECT IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, Estado " & _
        "FROM TbRiesgosMaterializaciones WHERE ID=" & matIdLong, dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Target materialization row not found for ID=" & matIdLong, logs)
        GoTo Teardown
    End If
    If CLng(rs.fields("IDProyecto").value) <> Test_Fixtures.Cache_ProyectoId Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("IDProyecto changed", logs)
        GoTo Teardown
    End If
    If CLng(rs.fields("IDEdicion").value) <> Test_Fixtures.Cache_EdicionId Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("IDEdicion changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("CodigoRiesgo").value, "")), codigoRiesgo, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("CodigoRiesgo changed", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("Fecha").value) <> targetFechaBefore Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Fecha changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("EsMaterializacion").value, "")), "Sí", vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("EsMaterializacion changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("Estado").value, "")), targetEstadoBefore, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Estado changed", logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    ' 14. Assert cardinalidad (VincularNC is EDIT-only, no insert).
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & TARGET_MAT_ID, dbOpenSnapshot)
    If rs.EOF Then
        rowsForTargetAfter = 0
    Else
        rowsForTargetAfter = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close: Set rs = Nothing
    If rowsForTargetAfter <> rowsForTargetBefore Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Cardinalidad: expected " & rowsForTargetBefore & " target rows, got " & _
                rowsForTargetAfter & " (VincularNC must be EDIT-only)", logs)
        GoTo Teardown
    End If

    ' 15. Assert unrelated guard row is preserved (its own decision fields untouched).
    Set rs = db.OpenRecordset( _
        "SELECT Fecha, Estado, IDNC, ParaNC, FechaDecison " & _
        "FROM TbRiesgosMaterializaciones WHERE ID=" & GUARD_MAT_ID, dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Unrelated guard row disappeared", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("Fecha").value) <> guardFechaBefore Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Guard Fecha changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("Estado").value, "")), guardEstadoBefore, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Guard Estado changed", logs)
        GoTo Teardown
    End If
    If CLng(Nz(rs.fields("IDNC").value, 0)) <> guardIDNCBefore Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Guard IDNC changed", logs)
        GoTo Teardown
    End If
    If StrComp(CStr(Nz(rs.fields("ParaNC").value, "")), guardParaNCBefore, vbTextCompare) <> 0 Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Guard ParaNC changed", logs)
        GoTo Teardown
    End If
    If CDate(rs.fields("FechaDecison").value) <> guardFechaDecisionBefore Then
        Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
            BuildFail("Guard FechaDecison changed", logs)
        GoTo Teardown
    End If
    rs.Close: Set rs = Nothing

    Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
        BuildOk("vincular_nc_persiste_triple", logs)

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set mat = Nothing
    Set nc = Nothing
    If transStarted Then
        ws.Rollback
        transStarted = False
    Else
        ' Defensive fallback only if the transaction was not opened. Never touch
        ' TbNoConformidades here because it is a linked shared backend table.
        Set cleanupDb = Test_Fixtures.GetTestDb(cleanupErr)
        If Not cleanupDb Is Nothing Then
            cleanupDb.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID IN (" & _
                TARGET_MAT_ID & ", " & GUARD_MAT_ID & ")"
            Set cleanupDb = Nothing
        End If
    End If
    Set ws = Nothing
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
    Exit Function

EH:
    Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison = _
        BuildFail("Test_RiesgoMaterializacion_VincularNC_PersisteTripleIDNCParaNCFechaDecison: " & _
            Err.description, logs)
    Resume Teardown
End Function

' ============================================================
' F.5 sad path (issue #34, PR 6) — VincularNC rejects when NC is Nothing
' Inverse precondition of F.5 happy path: target materialization
' exists with NC triple Null, NO NC row in TbNoConformidades, and
' the caller passes Nothing as the NC argument. The contract being
' tested is "VincularNC must reject with errMsg when NC is Nothing,
' and must NOT mutate the row". Mirrors F.3 sad path
' (RechazaNCInexistente) shape because no linked-NC transaction
' is needed (we never touch TbNoConformidades in this test). The
' design.md note flagged this contract as deferred: the current
' production code raises error 91 ("object variable or With block
' variable not set") on `Me.IDNC = p_ObjNC.IDNoConformidad` and
' the errores: block catches it (Err.Number <> 1000) and sets
' p_Error to a leaky internal message. The test asserts ONLY that
' errMsg is non-empty and that the DB row is untouched, matching
' the F.3 sad path assertion shape and leaving the production fix
' to a follow-up slice.
' ============================================================
Public Function Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente() As String
    On Error GoTo EH

    Dim logs(0 To 6) As String
    Dim errMsg As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim mat As RiesgoMaterializacion
    Dim codigoRiesgo As String
    Dim migrationResult As EnumSiNo
    Dim matIdLong As Long
    Dim rowsForKeyBefore As Long
    Dim rowsForKeyAfter As Long
    Dim idncInRow As Variant
    Dim paraNcInRow As Variant
    Dim fechaDecisonInRow As Variant
    Dim cleanupDb As DAO.Database
    Dim cleanupErr As String
    Const TARGET_MAT_ID As Long = 910950

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: SeedAll fixture graph (project=900501, riesgo=900503)"
    logs(2) = "3. Arrange: ensure IDPlanContingencia column via MigracionMaterializacionPlanContingencia_AsegurarCampo"
    logs(3) = "4. Arrange: pre-existing 'Sí' materialization row at ID=" & TARGET_MAT_ID & " (NC triple Null, NO NC row in TbNoConformidades)"
    logs(4) = "5. Act: VincularNC Nothing, errMsg (caller passes Nothing as the NC argument)"
    logs(5) = "6. Assert: errMsg is non-empty AND target row triple unchanged (IDNC=Null, ParaNC=Null, FechaDecison=Null) AND cardinalidad 1 -> 1"
    logs(6) = "7. Teardown: delete pre-existing materialization row + TeardownAll"

    ' 1. ForceLocalBackend
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = BuildFail(errMsg, logs)
        Exit Function
    End If

    ' 2. SeedAll
    Test_Fixtures.SeedAll

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        GoTo Teardown
    End If

    ' 3. Ensure IDPlanContingencia column (idempotent safety net)
    migrationResult = MigracionMaterializacionPlanContingencia_AsegurarCampo(errMsg)
    If migrationResult <> EnumSiNo.Sí Then
        Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
            BuildFail("Migration must succeed: " & errMsg, logs)
        GoTo Teardown
    End If
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
            BuildFail("Migration must not return errMsg: " & errMsg, logs)
        GoTo Teardown
    End If

    ' 4. Resolve CodigoRiesgo from the fixture riesgo (do not hardcode)
    Set rs = db.OpenRecordset( _
        "SELECT CodigoRiesgo FROM TbRiesgos WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, _
        dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
            BuildFail("Fixture riesgo " & Test_Fixtures.Cache_RiesgoId & " not found", logs)
        GoTo Teardown
    End If
    codigoRiesgo = CStr(Nz(rs.fields("CodigoRiesgo").value, ""))
    rs.Close: Set rs = Nothing
    If codigoRiesgo = "" Then
        Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
            BuildFail("Fixture riesgo CodigoRiesgo is empty", logs)
        GoTo Teardown
    End If

    ' 5. Pre-existing 'Sí' materialization row with all three NC decision
    '    fields NULL. NO NC row inserted in TbNoConformidades: this test
    '    exercises the "VincularNC must reject when p_ObjNC Is Nothing"
    '    contract, not the "Constructor.getNC returns Nothing" path that
    '    F.3 sad path already covers for RegistrarParaNC. We do not need
    '    a BeginTrans because we never touch the linked NC table.
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID=" & TARGET_MAT_ID, dbFailOnError
    Err.Clear
    On Error GoTo EH

    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, Estado) " & _
        "VALUES (" & TARGET_MAT_ID & ", " & Test_Fixtures.Cache_ProyectoId & ", " & _
        Test_Fixtures.Cache_EdicionId & ", '" & Replace(codigoRiesgo, "'", "''") & "', " & _
        "Date(), 'Sí', 'Materializado')", dbFailOnError

    ' 6. Cardinalidad ANTES — should be exactly 1 materialization row for fixture key
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId & _
        " AND CodigoRiesgo='" & Replace(codigoRiesgo, "'", "''") & "'", _
        dbOpenSnapshot)
    If rs.EOF Then
        rowsForKeyBefore = 0
    Else
        rowsForKeyBefore = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close: Set rs = Nothing
    If rowsForKeyBefore <> 1 Then
        Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
            BuildFail("Cardinalidad inicial: expected exactly 1 materialization row for fixture key, got " & _
                rowsForKeyBefore, logs)
        GoTo Teardown
    End If

    ' 7. Load materialization via Constructor.getRiesgoMaterializado
    errMsg = ""
    Set mat = Constructor.getRiesgoMaterializado(CStr(TARGET_MAT_ID), errMsg)
    If errMsg <> "" Then
        Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
            BuildFail("getRiesgoMaterializado failed: " & errMsg, logs)
        GoTo Teardown
    End If
    If mat Is Nothing Then
        Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
            BuildFail("getRiesgoMaterializado returned Nothing for ID=" & TARGET_MAT_ID, logs)
        GoTo Teardown
    End If
    matIdLong = TARGET_MAT_ID

    ' 8. Act: VincularNC with Nothing as the NC argument. The contract is
    '    "VincularNC must reject with errMsg when p_ObjNC Is Nothing, and
    '    must NOT mutate the row". Today the production code raises
    '    error 91 on `Me.IDNC = p_ObjNC.IDNoConformidad`; the errores:
    '    block catches it (Err.Number <> 1000) and sets p_Error to a
    '    leaky internal message. The test asserts ONLY that errMsg is
    '    non-empty (mirroring the F.3 sad path assertion shape) and
    '    that the DB row is untouched, so a follow-up production fix
    '    that replaces the leaky message with a clean contract message
    '    will not break this test.
    errMsg = ""
    mat.VincularNC Nothing, errMsg
    ' Sad path: errMsg MUST be non-empty
    If errMsg = "" Then
        Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
            BuildFail("VincularNC should set errMsg when NC=Nothing, got empty errMsg", logs)
        GoTo Teardown
    End If

    ' 9. Assert target row triple is UNCHANGED (IDNC=Null, ParaNC=Null, FechaDecison=Null)
    Set rs = db.OpenRecordset( _
        "SELECT IDNC, ParaNC, FechaDecison FROM TbRiesgosMaterializaciones " & _
        "WHERE ID=" & matIdLong, dbOpenSnapshot)
    If rs.EOF Then
        Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
            BuildFail("Materialization row not found for ID=" & matIdLong, logs)
        GoTo Teardown
    End If

    idncInRow = rs.fields("IDNC").value
    paraNcInRow = rs.fields("ParaNC").value
    fechaDecisonInRow = rs.fields("FechaDecison").value
    rs.Close: Set rs = Nothing

    If Not IsNull(idncInRow) Then
        Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
            BuildFail("IDNC on row should remain Null after rejected VincularNC(Nothing), got " & idncInRow, logs)
        GoTo Teardown
    End If
    If Not IsNull(paraNcInRow) Then
        Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
            BuildFail("ParaNC on row should remain Null after rejected VincularNC(Nothing), got '" & paraNcInRow & "'", logs)
        GoTo Teardown
    End If
    If Not IsNull(fechaDecisonInRow) Then
        Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
            BuildFail("FechaDecison on row should remain Null after rejected VincularNC(Nothing), got '" & fechaDecisonInRow & "'", logs)
        GoTo Teardown
    End If

    ' 10. Assert cardinalidad (no new row, no edit to existing)
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM TbRiesgosMaterializaciones " & _
        "WHERE IDProyecto=" & Test_Fixtures.Cache_ProyectoId & _
        " AND CodigoRiesgo='" & Replace(codigoRiesgo, "'", "''") & "'", _
        dbOpenSnapshot)
    If rs.EOF Then
        rowsForKeyAfter = 0
    Else
        rowsForKeyAfter = CLng(Nz(rs.fields("Cnt").value, 0))
    End If
    rs.Close: Set rs = Nothing
    If rowsForKeyAfter <> 1 Then
        Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
            BuildFail("Cardinalidad final: expected exactly 1 materialization row for fixture key, got " & _
                rowsForKeyAfter, logs)
        GoTo Teardown
    End If
    If rowsForKeyAfter <> rowsForKeyBefore Then
        Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
            BuildFail("Cardinalidad: expected rowsForKey unchanged (" & rowsForKeyBefore & _
                "), got " & rowsForKeyAfter, logs)
        GoTo Teardown
    End If

    Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
        BuildOk("vincular_nc_rechaza_nc_inexistente", logs)

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set mat = Nothing
    Set cleanupDb = Test_Fixtures.GetTestDb(cleanupErr)
    If Not cleanupDb Is Nothing Then
        cleanupDb.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID=" & TARGET_MAT_ID
        Set cleanupDb = Nothing
    End If
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
    Exit Function

EH:
    Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente = _
        BuildFail("Test_RiesgoMaterializacion_VincularNC_RechazaNCInexistente: " & _
            Err.description, logs)
    Resume Teardown
End Function

