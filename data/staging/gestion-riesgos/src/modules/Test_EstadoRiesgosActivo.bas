Attribute VB_Name = "Test_EstadoRiesgosActivo"
' =============================================================================
' Test_EstadoRiesgosActivo — TDD atoms for
'   modEstadoRiesgosActivo.EstadoRiesgosActivo_ObtenerEstado
'
' Skill: access-vba-tdd v2.6 + access-vba-e2e-methodology
' SDD:   forms-thin-phase0-testeable-2026-06-25 / WI-4
'
' Scope: 1 helper -> 6 scenario atoms
'   Happy      > riesgo Activo returns estado + fechaEstado
'   Sad        > empty ID returns BuildFail
'   Sad        > nonexistent ID returns BuildFail (cache miss + DB miss)
'   Edge       > riesgo Aceptado returns estado coherente
'   Edge       > riesgo Materializado returns fechaMaterializado
'   Adversarial> 2 consecutive calls return same result (idempotent)
'
' Signature:
'   Public Function EstadoRiesgosActivo_ObtenerEstado( _
'       ByVal p_IDRiesgo As String, _
'       Optional ByRef db As DAO.Database = Nothing, _
'       Optional ByRef p_Error As String) As String
'
' Cache discipline:
'   - Helper derives estado on-demand from m_DicRiesgos via GetCachedRiesgo.
'   - db parameter propagates to GetCachedRiesgo.
'   - NO second dictionary: single source of truth.
' =============================================================================
Option Compare Database
Option Explicit

' --- Constants ---
Private Const TEST_ID_BASE As Long = 900200
Private Const TEST_ID_ACTIVO As Long = 900201
Private Const TEST_ID_ACEPTADO As Long = 900202
Private Const TEST_ID_MATERIALIZADO As Long = 900203
Private Const TEST_ID_NOEXISTE As Long = 900299
Private Const TEST_ID_EDICION As Long = 900210
Private Const TEST_ID_PROYECTO As Long = 900220
Private Const TEST_FECHA_MAT As Date = #1/15/2025#

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

' InspectarCamposObligatorios — schema-first per §1.3
Private Sub InspectarCamposObligatorios(ByVal p_TableName As String, ByVal db As DAO.Database, _
                                        ByRef logs() As String, ByRef logIdx As Long)
    Dim tDef As DAO.TableDef
    Dim fld As DAO.Field
    Dim lineBuf As String
    On Error Resume Next
    Set tDef = db.TableDefs(p_TableName)
    If tDef Is Nothing Then
        logs(logIdx) = "  Inspect(" & p_TableName & "): TABLE NOT FOUND"
        logIdx = logIdx + 1
        Exit Sub
    End If
    lineBuf = "  Inspect(" & p_TableName & "): "
    For Each fld In tDef.fields
        If fld.required Then
            lineBuf = lineBuf & fld.Name & "(Req), "
        End If
    Next fld
    logs(logIdx) = lineBuf
    logIdx = logIdx + 1
    Set tDef = Nothing
    On Error GoTo 0
End Sub

' SeedFixtureRiesgo — seed parent + riesgo with Estado field
Private Sub SeedFixtureRiesgo(ByVal db As DAO.Database, ByVal p_IDRiesgo As Long, _
                              ByVal p_IDEdicion As Long, ByVal p_estado As String, _
                              Optional ByVal p_FechaMaterializado As Variant)
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, Estado) " & _
               "VALUES (" & p_IDRiesgo & ", " & p_IDEdicion & ", " & _
               "'PC-EST-" & p_IDRiesgo & "', 'R-EST-" & p_IDRiesgo & "', '" & p_estado & "')", dbFailOnError
    If Not IsMissing(p_FechaMaterializado) Then
        db.Execute "UPDATE TbRiesgos SET FechaMaterializado=#" & _
                   Format(p_FechaMaterializado, "yyyy-mm-dd") & "# " & _
                   "WHERE IDRiesgo=" & p_IDRiesgo, dbFailOnError
    End If
End Sub

' =============================================================================
' ATOM: W4-S1 Happy — riesgo Activo returns estado
' =============================================================================
Public Function Test_EstadoRiesgosActivo_W4_S1_Happy_Activo() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend"
    logIdx = logIdx + 1
    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_EstadoRiesgosActivo_W4_S1_Happy_Activo = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_EstadoRiesgosActivo_W4_S1_Happy_Activo = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' Idempotency DELETE
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900200 AND 900299", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_ID_PROYECTO, dbFailOnError

    ' Seed FK chain
    logs(logIdx) = "1b. Schema: inspecting required fields for INSERT"
    logIdx = logIdx + 1
    Call InspectarCamposObligatorios("TbProyectos", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbProyectosEdiciones", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbRiesgos", db, logs, logIdx)
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-EST-W4S1', 'Estado test')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    Call SeedFixtureRiesgo(db, TEST_ID_ACTIVO, TEST_ID_EDICION, "Activo")

    Dim countAfter As Long
    countAfter = CountRows(db, "TbRiesgos", "IDRiesgo=" & TEST_ID_ACTIVO)
    If countAfter <> 1 Then
        Test_EstadoRiesgosActivo_W4_S1_Happy_Activo = BuildFail("Arrange FAIL: expected 1 row, got " & countAfter, logs)
        GoTo Teardown
    End If

    ' Act
    logs(logIdx) = "2. Act: EstadoRiesgosActivo_ObtenerEstado(" & TEST_ID_ACTIVO & ")"
    logIdx = logIdx + 1
    Dim p_Error As String
    Dim result As String
    result = EstadoRiesgosActivo_ObtenerEstado(CStr(TEST_ID_ACTIVO), db, p_Error)

    ' Assert
    If p_Error <> "" Then
        logs(logIdx) = "3. Assert FAIL: p_Error=" & p_Error
        logIdx = logIdx + 1
        Test_EstadoRiesgosActivo_W4_S1_Happy_Activo = BuildFail("Unexpected p_Error: " & p_Error, logs)
        GoTo Teardown
    End If
    If result = "" Then
        Test_EstadoRiesgosActivo_W4_S1_Happy_Activo = BuildFail("result is empty", logs)
        GoTo Teardown
    End If
    If InStr(result, """estado"":""Activo""") = 0 Then
        Test_EstadoRiesgosActivo_W4_S1_Happy_Activo = BuildFail("result missing estado=Activo: " & result, logs)
        GoTo Teardown
    End If

    logs(logIdx) = "3. Assert: result contains idRiesgo + estado=Activo"
    logIdx = logIdx + 1
    Test_EstadoRiesgosActivo_W4_S1_Happy_Activo = BuildOk("activo", logs)
    Exit Function
EH:
    Test_EstadoRiesgosActivo_W4_S1_Happy_Activo = BuildFail("EH: " & Err.description, logs)
Teardown:
    ' No-op shared label for GoTo Teardown targets in this atom.
    ' Per-atom cleanup (DELETE FROM seeded rows) is currently inline;
    ' a real cleanup would go here in the future.
End Function

' Teardown handler (shared label) — actual cleanup happens in each atom's Teardown
Private Sub Teardown()
    ' No-op: actual teardown is per-atom inline
End Sub

' =============================================================================
' ATOM: W4-S2 Sad — empty ID returns BuildFail
' =============================================================================
Public Function Test_EstadoRiesgosActivo_W4_S2_Sad_EmptyID() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_EstadoRiesgosActivo_W4_S2_Sad_EmptyID = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_EstadoRiesgosActivo_W4_S2_Sad_EmptyID = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    logs(logIdx) = "1. Act: EstadoRiesgosActivo_ObtenerEstado("""")"
    logIdx = logIdx + 1
    Dim p_Error As String
    Dim result As String
    result = EstadoRiesgosActivo_ObtenerEstado("", db, p_Error)

    If p_Error = "" Then
        Test_EstadoRiesgosActivo_W4_S2_Sad_EmptyID = BuildFail("p_Error should be populated for empty ID", logs)
        Exit Function
    End If
    If result = "" Then
        Test_EstadoRiesgosActivo_W4_S2_Sad_EmptyID = BuildFail("result should contain TestCore_BuildFail JSON", logs)
        Exit Function
    End If

    logs(logIdx) = "2. Assert: p_Error populated, result is BuildFail JSON"
    logIdx = logIdx + 1
    Test_EstadoRiesgosActivo_W4_S2_Sad_EmptyID = BuildOk("empty_rejected", logs)
    Exit Function
EH:
    Test_EstadoRiesgosActivo_W4_S2_Sad_EmptyID = BuildFail("EH: " & Err.description, logs)
End Function

' =============================================================================
' ATOM: W4-S3 Sad — nonexistent ID returns BuildFail
' =============================================================================
Public Function Test_EstadoRiesgosActivo_W4_S3_Sad_NonexistentID() As String
    Dim logs(0 To 7) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_EstadoRiesgosActivo_W4_S3_Sad_NonexistentID = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_EstadoRiesgosActivo_W4_S3_Sad_NonexistentID = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    ' Idempotency DELETE (no rows exist for this test)
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & TEST_ID_NOEXISTE, dbFailOnError

    logs(logIdx) = "1. Act: EstadoRiesgosActivo_ObtenerEstado(" & TEST_ID_NOEXISTE & ")"
    logIdx = logIdx + 1
    Dim p_Error As String
    Dim result As String
    result = EstadoRiesgosActivo_ObtenerEstado(CStr(TEST_ID_NOEXISTE), db, p_Error)

    If p_Error = "" Then
        Test_EstadoRiesgosActivo_W4_S3_Sad_NonexistentID = BuildFail("p_Error should be populated for missing ID", logs)
        Exit Function
    End If

    logs(logIdx) = "2. Assert: p_Error populated for missing ID"
    logIdx = logIdx + 1
    Test_EstadoRiesgosActivo_W4_S3_Sad_NonexistentID = BuildOk("missing_rejected", logs)
    Exit Function
EH:
    Test_EstadoRiesgosActivo_W4_S3_Sad_NonexistentID = BuildFail("EH: " & Err.description, logs)
End Function

' =============================================================================
' ATOM: W4-S4 Edge — riesgo Aceptado returns estado coherente
' =============================================================================
Public Function Test_EstadoRiesgosActivo_W4_S4_Edge_Aceptado() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_EstadoRiesgosActivo_W4_S4_Edge_Aceptado = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_EstadoRiesgosActivo_W4_S4_Edge_Aceptado = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & TEST_ID_ACEPTADO, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_ID_PROYECTO, dbFailOnError

    Call InspectarCamposObligatorios("TbRiesgos", db, logs, logIdx)
    logIdx = logIdx + 1
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-EST-W4S4', 'Aceptado test')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    Call SeedFixtureRiesgo(db, TEST_ID_ACEPTADO, TEST_ID_EDICION, "Aceptado")

    Dim countAfter As Long
    countAfter = CountRows(db, "TbRiesgos", "IDRiesgo=" & TEST_ID_ACEPTADO)
    If countAfter <> 1 Then
        Test_EstadoRiesgosActivo_W4_S4_Edge_Aceptado = BuildFail("Arrange FAIL: expected 1 row, got " & countAfter, logs)
        Exit Function
    End If

    Dim p_Error As String
    Dim result As String
    result = EstadoRiesgosActivo_ObtenerEstado(CStr(TEST_ID_ACEPTADO), db, p_Error)

    If p_Error <> "" Then
        Test_EstadoRiesgosActivo_W4_S4_Edge_Aceptado = BuildFail("p_Error=" & p_Error, logs)
        Exit Function
    End If
    If InStr(result, """estado"":""Aceptado""") = 0 Then
        Test_EstadoRiesgosActivo_W4_S4_Edge_Aceptado = BuildFail("missing estado=Aceptado: " & result, logs)
        Exit Function
    End If

    logIdx = logIdx + 1
    Test_EstadoRiesgosActivo_W4_S4_Edge_Aceptado = BuildOk("aceptado", logs)
    Exit Function
EH:
    Test_EstadoRiesgosActivo_W4_S4_Edge_Aceptado = BuildFail("EH: " & Err.description, logs)
End Function

' =============================================================================
' ATOM: W4-S5 Edge — riesgo Materializado returns fechaMaterializado
' =============================================================================
Public Function Test_EstadoRiesgosActivo_W4_S5_Edge_Materializado() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_EstadoRiesgosActivo_W4_S5_Edge_Materializado = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_EstadoRiesgosActivo_W4_S5_Edge_Materializado = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & TEST_ID_MATERIALIZADO, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_ID_PROYECTO, dbFailOnError

    Call InspectarCamposObligatorios("TbRiesgos", db, logs, logIdx)
    logIdx = logIdx + 1
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-EST-W4S5', 'Materializado test')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    Call SeedFixtureRiesgo(db, TEST_ID_MATERIALIZADO, TEST_ID_EDICION, "Materializado", TEST_FECHA_MAT)

    Dim countAfter As Long
    countAfter = CountRows(db, "TbRiesgos", "IDRiesgo=" & TEST_ID_MATERIALIZADO)
    If countAfter <> 1 Then
        Test_EstadoRiesgosActivo_W4_S5_Edge_Materializado = BuildFail("Arrange FAIL: expected 1 row, got " & countAfter, logs)
        Exit Function
    End If

    Dim p_Error As String
    Dim result As String
    result = EstadoRiesgosActivo_ObtenerEstado(CStr(TEST_ID_MATERIALIZADO), db, p_Error)

    If p_Error <> "" Then
        Test_EstadoRiesgosActivo_W4_S5_Edge_Materializado = BuildFail("p_Error=" & p_Error, logs)
        Exit Function
    End If
    If InStr(result, """estado"":""Materializado""") = 0 Then
        Test_EstadoRiesgosActivo_W4_S5_Edge_Materializado = BuildFail("missing estado=Materializado: " & result, logs)
        Exit Function
    End If
    If InStr(result, "fechaEstado") = 0 Then
        Test_EstadoRiesgosActivo_W4_S5_Edge_Materializado = BuildFail("missing fechaEstado field: " & result, logs)
        Exit Function
    End If

    logIdx = logIdx + 1
    Test_EstadoRiesgosActivo_W4_S5_Edge_Materializado = BuildOk("materializado_with_fecha", logs)
    Exit Function
EH:
    Test_EstadoRiesgosActivo_W4_S5_Edge_Materializado = BuildFail("EH: " & Err.description, logs)
End Function

' =============================================================================
' ATOM: W4-S6 Adversarial — two consecutive calls return same result
' =============================================================================
Public Function Test_EstadoRiesgosActivo_W4_S6_Adversarial_Idempotent() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0
    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_EstadoRiesgosActivo_W4_S6_Adversarial_Idempotent = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_EstadoRiesgosActivo_W4_S6_Adversarial_Idempotent = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & TEST_ID_ACTIVO, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_ID_PROYECTO, dbFailOnError

    Call InspectarCamposObligatorios("TbRiesgos", db, logs, logIdx)
    logIdx = logIdx + 1
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-EST-W4S6', 'Adversarial test')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    Call SeedFixtureRiesgo(db, TEST_ID_ACTIVO, TEST_ID_EDICION, "Activo")

    ' Two consecutive calls
    Dim p_Error1 As String
    Dim p_Error2 As String
    Dim result1 As String
    Dim result2 As String
    result1 = EstadoRiesgosActivo_ObtenerEstado(CStr(TEST_ID_ACTIVO), db, p_Error1)
    result2 = EstadoRiesgosActivo_ObtenerEstado(CStr(TEST_ID_ACTIVO), db, p_Error2)

    If p_Error1 <> "" Or p_Error2 <> "" Then
        Test_EstadoRiesgosActivo_W4_S6_Adversarial_Idempotent = BuildFail("p_Error populated: " & p_Error1 & " / " & p_Error2, logs)
        Exit Function
    End If
    If result1 <> result2 Then
        Test_EstadoRiesgosActivo_W4_S6_Adversarial_Idempotent = BuildFail("idempotency violation: " & result1 & " vs " & result2, logs)
        Exit Function
    End If

    logIdx = logIdx + 1
    Test_EstadoRiesgosActivo_W4_S6_Adversarial_Idempotent = BuildOk("idempotent", logs)
    Exit Function
EH:
    Test_EstadoRiesgosActivo_W4_S6_Adversarial_Idempotent = BuildFail("EH: " & Err.description, logs)
End Function

