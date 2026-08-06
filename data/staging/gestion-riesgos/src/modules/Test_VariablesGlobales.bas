Attribute VB_Name = "Test_VariablesGlobales"
' ============================================================
' Test_VariablesGlobales — TDD Red Atoms for
'   Variables_Globales.GetCachedRiesgoFresh (NEW, Block 1B WI-1)
'
' Skill: access-vba-tdd v2.5 + access-vba-e2e-methodology
'
' Scope: 1 helper -> 6 scenario classes
'   Happy      > cold cache miss returns new riesgo
'   Happy      > cached instance is invalidated and replaced (B <> A)
'   Sad        > empty/whitespace/Null ID returns Nothing + p_Error
'   Sad        > non-existent ID returns Nothing + p_Error
'   Edge       > charset / padded / special-char IDs are deterministic
'   Adversarial> double consecutive call returns distinct instances
'
' Helper signature (canonica, ADJACENT to GetCachedRiesgo in Variables Globales.bas):
'   Public Function GetCachedRiesgoFresh( _
'       ByVal p_IDRiesgo As String, _
'       Optional ByRef p_Error As String) As riesgo
'
' Notes:
'   - Guarantees an instance DISTINCT from the previous cache contents.
'     Implemented as: InvalidarCacheRiesgo + GetCachedRiesgo. Without this
'     pattern, two consecutive GetCachedRiesgo calls return the SAME instance
'     (cache hit on the same reference), breaking the diff baseline in
'     RiesgoChangeDetector.HasChanges (see Form_FormRiesgo.cls:266 latent bug).
'   - p_Error convention (Telefonica D&S): empty string on success;
'     populated on real errors (empty ID, ID not found, DAO failure).
'   - This module references GetCachedRiesgoFresh BEFORE the helper exists.
'     That is intentional: when the user compiles the project in VBE
'     (Debug -> Compile), VBE will report "Sub or Function not defined"
'     for GetCachedRiesgoFresh -> RED state. Once the helper is added
'     (Commit 1.2), the module compiles and the atoms go GREEN.
' ============================================================
Option Compare Database
Option Explicit

' ============================================================
' JSON wrappers (BuildOk/BuildFail) — must precede all Private Sub/Function
' Per access-vba-tdd skill §1.8
' ============================================================
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' ============================================================
' Private helpers
' ============================================================

Private Sub InspectarCamposObligatorios(ByVal p_TableName As String, ByVal db As DAO.Database, _
                                        ByRef logs() As String, ByRef logIdx As Long)
    ' Schema inspection: recorre campos Required/AllowZeroLength de la tabla
    ' y loguea hallazgos antes del primer INSERT del atomo.
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
    For Each fld In tDef.Fields
        If fld.Required Then
            lineBuf = lineBuf & fld.Name & "(Req), "
        ElseIf fld.Type = dbText And Not fld.Required Then
            ' Only log if AllowZeroLength is relevant
        End If
    Next fld
    logs(logIdx) = lineBuf
    logIdx = logIdx + 1
    Set tDef = Nothing
    On Error GoTo 0
End Sub

' ============================================================
' CountRows — cardinality helper per access-vba-tdd §4.5
' Pattern: every Test_*.bas has its own Private CountRows (Test_IndicadorCobertura,
' Test_RiesgoEstadoGateHelper, etc.). Adding here for consistency with that pattern.
' ============================================================
Private Function CountRows(ByVal db As DAO.Database, ByVal p_Table As String, _
                           ByVal p_Where As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS c FROM " & p_Table & " WHERE " & p_Where)
    CountRows = CLng(Nz(rs!c, 0))
    rs.Close
    Set rs = Nothing
End Function

' ============================================================
' ATOM: W1-S1 Happy cold
'   cache miss; valid ID; returns New riesgo (instance B); p_Error=""
'
' Setup: seed riesgo IDRiesgo=900001 (cold cache — no prior GetCachedRiesgo)
' Act:   GetCachedRiesgoFresh("900001", p_Error)
' Expect: result <> Nothing, p_Error="", result is a riesgo instance
' Teardown: DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999
' ============================================================
Public Function Test_GetCachedRiesgoFresh_W1_S1_Happy_Cold() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend"
    logIdx = logIdx + 1

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_GetCachedRiesgoFresh_W1_S1_Happy_Cold = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_GetCachedRiesgoFresh_W1_S1_Happy_Cold = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Const TEST_ID_EDICION As Long = 900070
    Const TEST_ID_PROYECTO As Long = 900080
    Const TEST_ID_RIESGO As Long = 900001

    ' Clean any residual from prior runs (FK order: child -> parent)
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion BETWEEN 900070 AND 900079", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto BETWEEN 900080 AND 900089", dbFailOnError

    ' Seed FK chain: Proyecto -> Edicion -> Riesgo
    logs(logIdx) = "1b. Schema: inspecting required fields for INSERT": logIdx = logIdx + 1
    Call InspectarCamposObligatorios("TbProyectos", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbProyectosEdiciones", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbRiesgos", db, logs, logIdx)

    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-CF-W1S1', 'Cold cache project')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) " & _
               "VALUES (" & TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", " & _
               "'PC-CF-W1S1', 'R-CF-W1S1')", dbFailOnError

    ' Cardinality per access-vba-tdd §4.5: verify mutation actually persisted.
    Dim countAfter As Long
    countAfter = CountRows(db, "TbRiesgos", "IDRiesgo=" & TEST_ID_RIESGO)
    If countAfter <> 1 Then
        Test_GetCachedRiesgoFresh_W1_S1_Happy_Cold = BuildFail( _
            "Arrange FAIL: expected 1 row seeded in TbRiesgos for IDRiesgo=" & TEST_ID_RIESGO & ", got " & countAfter, logs)
        GoTo Teardown
    End If

    logs(logIdx) = "2. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO & " (cold cache, countAfter=" & countAfter & ")"
    logIdx = logIdx + 1

    ' -- Act -------------------------------------------------
    logs(logIdx) = "3. Act: GetCachedRiesgoFresh(" & TEST_ID_RIESGO & ")"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim result As riesgo
    Set result = GetCachedRiesgoFresh(CStr(TEST_ID_RIESGO), p_Error, db)

    ' -- Assert ----------------------------------------------
    If p_Error <> "" Then
        logs(logIdx) = "4. Assert FAIL: p_Error populated=" & p_Error
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S1_Happy_Cold = BuildFail( _
            "Unexpected error: " & p_Error, logs)
        GoTo Teardown
    End If

    If result Is Nothing Then
        logs(logIdx) = "4. Assert FAIL: result is Nothing"
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S1_Happy_Cold = BuildFail( _
            "GetCachedRiesgoFresh must return a riesgo instance on cold cache miss", logs)
        GoTo Teardown
    End If

    logs(logIdx) = "4. Assert PASS: returned riesgo instance for ID=" & TEST_ID_RIESGO
    logIdx = logIdx + 1
    Test_GetCachedRiesgoFresh_W1_S1_Happy_Cold = BuildOk("cold_cache_hit", logs)
    GoTo Teardown

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_GetCachedRiesgoFresh_W1_S1_Happy_Cold = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999", dbFailOnError
    On Error GoTo 0
End Function

' ============================================================
' ATOM: W1-S2 Happy cached
'   cache hit with instance A; GetCachedRiesgoFresh invalidates and returns B (B <> A)
'
' Setup: seed riesgo IDRiesgo=900002
' Act:   1) GetCachedRiesgo("900002") -> A
'        2) GetCachedRiesgoFresh("900002") -> B
' Expect: A <> Nothing, B <> Nothing, B Is NOT A (distinct instances)
' ============================================================
Public Function Test_GetCachedRiesgoFresh_W1_S2_Happy_Cached() As String
    Dim logs(0 To 12) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_GetCachedRiesgoFresh_W1_S2_Happy_Cached = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_GetCachedRiesgoFresh_W1_S2_Happy_Cached = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Const TEST_ID_EDICION As Long = 900071
    Const TEST_ID_PROYECTO As Long = 900081
    Const TEST_ID_RIESGO As Long = 900002

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion BETWEEN 900070 AND 900079", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto BETWEEN 900080 AND 900089", dbFailOnError

    logs(logIdx) = "1b. Schema: inspecting required fields for INSERT": logIdx = logIdx + 1
    Call InspectarCamposObligatorios("TbProyectos", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbProyectosEdiciones", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbRiesgos", db, logs, logIdx)

    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-CF-W1S2', 'Cached project')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) " & _
               "VALUES (" & TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", " & _
               "'PC-CF-W1S2', 'R-CF-W1S2')", dbFailOnError

    Dim countAfter_S2 As Long
    countAfter_S2 = CountRows(db, "TbRiesgos", "IDRiesgo=" & TEST_ID_RIESGO)
    If countAfter_S2 <> 1 Then
        Test_GetCachedRiesgoFresh_W1_S2_Happy_Cached = BuildFail( _
            "Arrange FAIL: expected 1 row in TbRiesgos, got " & countAfter_S2, logs)
        GoTo Teardown
    End If

    logs(logIdx) = "1. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO & " (countAfter=" & countAfter_S2 & ")"
    logIdx = logIdx + 1

    ' -- Act: first populate cache with A
    logs(logIdx) = "2a. Act: GetCachedRiesgo(" & TEST_ID_RIESGO & ") -> A (cache hit afterwards)"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim instanceA As riesgo
    Set instanceA = GetCachedRiesgo(CStr(TEST_ID_RIESGO), p_Error)
    If p_Error <> "" Or instanceA Is Nothing Then
        logs(logIdx) = "3a. Assert FAIL: warmup failed: p_Error=" & p_Error
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S2_Happy_Cached = BuildFail( _
            "warmup GetCachedRiesgo failed", logs)
        GoTo Teardown
    End If

    ' -- Act: now Fresh -> should return B (distinct from A)
    logs(logIdx) = "2b. Act: GetCachedRiesgoFresh(" & TEST_ID_RIESGO & ") -> B (must be distinct from A)"
    logIdx = logIdx + 1

    Dim instanceB As riesgo
    p_Error = ""
    Set instanceB = GetCachedRiesgoFresh(CStr(TEST_ID_RIESGO), p_Error, db)

    ' -- Assert
    If p_Error <> "" Then
        logs(logIdx) = "3b. Assert FAIL: p_Error populated=" & p_Error
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S2_Happy_Cached = BuildFail( _
            "Unexpected error from GetCachedRiesgoFresh: " & p_Error, logs)
        GoTo Teardown
    End If

    If instanceB Is Nothing Then
        logs(logIdx) = "3b. Assert FAIL: instanceB is Nothing"
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S2_Happy_Cached = BuildFail( _
            "GetCachedRiesgoFresh must return a riesgo instance", logs)
        GoTo Teardown
    End If

    ' --- The critical assertion: B must NOT be A. If B Is A, the cache
    '     was not invalidated and GetCachedRiesgoFresh is broken (would
    '     break Form_FormRiesgo.cls:266 diff baseline).
    If instanceB Is instanceA Then
        logs(logIdx) = "3b. Assert FAIL: instanceB IS instanceA (same reference) — invalidation did not happen"
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S2_Happy_Cached = BuildFail( _
            "GetCachedRiesgoFresh must return a DISTINCT instance from the cached one", logs)
        GoTo Teardown
    End If

    logs(logIdx) = "3b. Assert PASS: A and B are DISTINCT instances (invalidation worked)"
    logIdx = logIdx + 1
    Test_GetCachedRiesgoFresh_W1_S2_Happy_Cached = BuildOk("fresh_instance_distinct_from_cache", logs)
    GoTo Teardown

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_GetCachedRiesgoFresh_W1_S2_Happy_Cached = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999", dbFailOnError
    On Error GoTo 0
End Function

' ============================================================
' ATOM: W1-S3 Sad empty
'   ID "" / "   " returns Nothing; p_Error mentions "vacio"
'
' Setup: no seed rows
' Act:   GetCachedRiesgoFresh("") and GetCachedRiesgoFresh("   ")
' Expect: both return Nothing; p_Error mentions empty / vacio
' ============================================================
Public Function Test_GetCachedRiesgoFresh_W1_S3_Sad_Empty() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_GetCachedRiesgoFresh_W1_S3_Sad_Empty = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    logs(logIdx) = "1. Arrange: ForceLocalBackend OK; no seed rows"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim result As riesgo

    ' -- Empty string
    logs(logIdx) = "2a. Act: GetCachedRiesgoFresh("""")"
    logIdx = logIdx + 1
    Set result = GetCachedRiesgoFresh("", p_Error)

    If Not result Is Nothing Then
        logs(logIdx) = "3a. Assert FAIL: expected Nothing, got instance"
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S3_Sad_Empty = BuildFail( _
            "Empty ID must return Nothing", logs)
        Exit Function
    End If
    If p_Error = "" Then
        logs(logIdx) = "3a. Assert FAIL: expected p_Error populated, got empty"
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S3_Sad_Empty = BuildFail( _
            "Empty ID must populate p_Error", logs)
        Exit Function
    End If
    If InStr(1, p_Error, "vacio", vbTextCompare) = 0 Then
        logs(logIdx) = "3a. Assert FAIL: p_Error does not mention 'vacio': " & p_Error
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S3_Sad_Empty = BuildFail( _
            "p_Error should describe empty input", logs)
        Exit Function
    End If
    logs(logIdx) = "3a. Assert PASS: empty ID rejected with p_Error=" & p_Error
    logIdx = logIdx + 1

    ' -- Whitespace-only
    logs(logIdx) = "2b. Act: GetCachedRiesgoFresh(""   "")"
    logIdx = logIdx + 1
    p_Error = ""
    Set result = GetCachedRiesgoFresh("   ", p_Error)

    If Not result Is Nothing Then
        logs(logIdx) = "3b. Assert FAIL: whitespace ID returned instance"
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S3_Sad_Empty = BuildFail( _
            "Whitespace ID must return Nothing", logs)
        Exit Function
    End If
    If p_Error = "" Then
        logs(logIdx) = "3b. Assert FAIL: whitespace ID returned no p_Error"
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S3_Sad_Empty = BuildFail( _
            "Whitespace ID must populate p_Error", logs)
        Exit Function
    End If

    logs(logIdx) = "3b. Assert PASS: whitespace ID rejected with p_Error=" & p_Error
    logIdx = logIdx + 1
    Test_GetCachedRiesgoFresh_W1_S3_Sad_Empty = BuildOk("empty_and_whitespace_rejected", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_GetCachedRiesgoFresh_W1_S3_Sad_Empty = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM: W1-S4 Sad error
'   non-existent ID returns Nothing; p_Error populated
'
' Setup: no seed rows for ID=900099
' Act:   GetCachedRiesgoFresh("900099", p_Error)
' Expect: result = Nothing; p_Error <> ""
' ============================================================
Public Function Test_GetCachedRiesgoFresh_W1_S4_Sad_NotFound() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_GetCachedRiesgoFresh_W1_S4_Sad_NotFound = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_GetCachedRiesgoFresh_W1_S4_Sad_NotFound = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999", dbFailOnError

    logs(logIdx) = "1. Arrange: ForceLocalBackend OK; ID=900099 missing"
    logIdx = logIdx + 1

    Const TEST_ID_MISSING As Long = 900099
    logs(logIdx) = "2. Act: GetCachedRiesgoFresh(" & TEST_ID_MISSING & ")"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim result As riesgo
    Set result = GetCachedRiesgoFresh(CStr(TEST_ID_MISSING), p_Error, db)

    If Not result Is Nothing Then
        logs(logIdx) = "3. Assert FAIL: expected Nothing, got instance"
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S4_Sad_NotFound = BuildFail( _
            "Non-existent ID must return Nothing", logs)
        Exit Function
    End If
    If p_Error = "" Then
        logs(logIdx) = "3. Assert FAIL: expected p_Error populated, got empty"
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S4_Sad_NotFound = BuildFail( _
            "Non-existent ID must populate p_Error", logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: non-existent ID rejected with p_Error=" & p_Error
    logIdx = logIdx + 1
    Test_GetCachedRiesgoFresh_W1_S4_Sad_NotFound = BuildOk("missing_id_rejected", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_GetCachedRiesgoFresh_W1_S4_Sad_NotFound = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM: W1-S5 Edge charset
'   Padded / special-char IDs are handled deterministically
'
' Setup: seed riesgo IDRiesgo=900003 (bare)
' Act:   GetCachedRiesgoFresh("  900003  ") -> must resolve to bare ID
' Expect: result <> Nothing (Access SQL leniency matches padded to bare)
'         OR Nothing + clear p_Error (whichever the helper chose; deterministic)
' ============================================================
Public Function Test_GetCachedRiesgoFresh_W1_S5_Edge_Charset() As String
    Dim logs(0 To 12) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_GetCachedRiesgoFresh_W1_S5_Edge_Charset = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_GetCachedRiesgoFresh_W1_S5_Edge_Charset = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Const TEST_ID_EDICION As Long = 900072
    Const TEST_ID_PROYECTO As Long = 900082
    Const TEST_ID_RIESGO As Long = 900003

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion BETWEEN 900070 AND 900079", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto BETWEEN 900080 AND 900089", dbFailOnError

    logs(logIdx) = "1b. Schema: inspecting required fields for INSERT": logIdx = logIdx + 1
    Call InspectarCamposObligatorios("TbProyectos", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbProyectosEdiciones", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbRiesgos", db, logs, logIdx)

    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-CF-W1S5', 'Edge charset project')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) " & _
               "VALUES (" & TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", " & _
               "'PC-CF-W1S5', 'R-CF-W1S5')", dbFailOnError

    Dim countAfter_S5 As Long
    countAfter_S5 = CountRows(db, "TbRiesgos", "IDRiesgo=" & TEST_ID_RIESGO)
    If countAfter_S5 <> 1 Then
        Test_GetCachedRiesgoFresh_W1_S5_Edge_Charset = BuildFail( _
            "Arrange FAIL: expected 1 row in TbRiesgos, got " & countAfter_S5, logs)
        GoTo Teardown
    End If

    logs(logIdx) = "1. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO & " (bare, countAfter=" & countAfter_S5 & ")"
    logIdx = logIdx + 1

    Dim paddedId As String
    paddedId = "  " & TEST_ID_RIESGO & "  "
    logs(logIdx) = "2. Act: GetCachedRiesgoFresh(""" & paddedId & """)"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim result As riesgo
    Set result = GetCachedRiesgoFresh(paddedId, p_Error, db)

    ' --- Deterministic expectation: Access SQL coerces padded ID to bare
    '     ID via leniency (same leniency documented in Test_RiesgoDetalleRefrescoHelper).
    '     We accept BOTH outcomes but require consistency:
    '       a) result <> Nothing -> resolved via SQL leniency
    '       b) result = Nothing  -> rejected, p_Error populated
    If result Is Nothing And p_Error = "" Then
        logs(logIdx) = "3. Assert FAIL: helper returned Nothing with no error message"
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S5_Edge_Charset = BuildFail( _
            "Helper must either resolve padded ID or populate p_Error", logs)
        GoTo Teardown
    End If

    If result Is Nothing Then
        logs(logIdx) = "3. Assert PASS: helper rejected padded ID with p_Error=" & p_Error
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S5_Edge_Charset = BuildOk("padded_id_rejected_deterministic", logs)
        GoTo Teardown
    End If

    logs(logIdx) = "3. Assert PASS: helper resolved padded ID via SQL leniency"
    logIdx = logIdx + 1
    Test_GetCachedRiesgoFresh_W1_S5_Edge_Charset = BuildOk("padded_id_resolved_via_sql_leniency", logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999", dbFailOnError
    On Error GoTo 0
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_GetCachedRiesgoFresh_W1_S5_Edge_Charset = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM: W1-S6 Adversarial concurrent
'   Double consecutive call with same ID returns distinct instances
'   (sanity check of the contract: VBA single-threaded but two calls
'    in series still produce different memory addresses for the
'    riesgo objects, because each Fresh invalidates the cache before
'    the second GetCachedRiesgo).
'
' Setup: seed riesgo IDRiesgo=900004
' Act:   Fresh #1 -> A; Fresh #2 -> B
' Expect: A <> Nothing, B <> Nothing, A Is NOT B (distinct memory refs)
' ============================================================
Public Function Test_GetCachedRiesgoFresh_W1_S6_Adversarial_Concurrent() As String
    Dim logs(0 To 12) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_GetCachedRiesgoFresh_W1_S6_Adversarial_Concurrent = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_GetCachedRiesgoFresh_W1_S6_Adversarial_Concurrent = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Const TEST_ID_EDICION As Long = 900073
    Const TEST_ID_PROYECTO As Long = 900083
    Const TEST_ID_RIESGO As Long = 900004

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion BETWEEN 900070 AND 900079", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto BETWEEN 900080 AND 900089", dbFailOnError

    logs(logIdx) = "1b. Schema: inspecting required fields for INSERT": logIdx = logIdx + 1
    Call InspectarCamposObligatorios("TbProyectos", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbProyectosEdiciones", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbRiesgos", db, logs, logIdx)

    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-CF-W1S6', 'Adversarial project')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) " & _
               "VALUES (" & TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", " & _
               "'PC-CF-W1S6', 'R-CF-W1S6')", dbFailOnError

    Dim countAfter_S6 As Long
    countAfter_S6 = CountRows(db, "TbRiesgos", "IDRiesgo=" & TEST_ID_RIESGO)
    If countAfter_S6 <> 1 Then
        Test_GetCachedRiesgoFresh_W1_S6_Adversarial_Concurrent = BuildFail( _
            "Arrange FAIL: expected 1 row in TbRiesgos, got " & countAfter_S6, logs)
        GoTo Teardown
    End If

    logs(logIdx) = "1. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO & " (countAfter=" & countAfter_S6 & ")"
    logIdx = logIdx + 1

    ' -- First Fresh call -> A
    logs(logIdx) = "2a. Act: GetCachedRiesgoFresh #1 -> A"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim instanceA As riesgo
    Set instanceA = GetCachedRiesgoFresh(CStr(TEST_ID_RIESGO), p_Error, db)
    If p_Error <> "" Or instanceA Is Nothing Then
        logs(logIdx) = "3a. Assert FAIL: Fresh #1 failed (err='" & p_Error & "')"
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S6_Adversarial_Concurrent = BuildFail( _
            "Fresh #1 must succeed", logs)
        GoTo Teardown
    End If

    ' -- Second Fresh call -> B (must be a fresh instance, not the cached A)
    logs(logIdx) = "2b. Act: GetCachedRiesgoFresh #2 -> B (must be fresh)"
    logIdx = logIdx + 1

    Dim instanceB As riesgo
    p_Error = ""
    Set instanceB = GetCachedRiesgoFresh(CStr(TEST_ID_RIESGO), p_Error, db)
    If p_Error <> "" Or instanceB Is Nothing Then
        logs(logIdx) = "3b. Assert FAIL: Fresh #2 failed (err='" & p_Error & "')"
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S6_Adversarial_Concurrent = BuildFail( _
            "Fresh #2 must succeed", logs)
        GoTo Teardown
    End If

    ' -- Critical: A and B must be DISTINCT instances
    If instanceB Is instanceA Then
        logs(logIdx) = "3b. Assert FAIL: B IS A — invalidation did not happen between the two calls"
        logIdx = logIdx + 1
        Test_GetCachedRiesgoFresh_W1_S6_Adversarial_Concurrent = BuildFail( _
            "Consecutive Fresh calls must return distinct instances", logs)
        GoTo Teardown
    End If

    logs(logIdx) = "3b. Assert PASS: A and B are distinct instances (concurrent-safe contract)"
    logIdx = logIdx + 1
    Test_GetCachedRiesgoFresh_W1_S6_Adversarial_Concurrent = BuildOk("concurrent_distinct", logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999", dbFailOnError
    On Error GoTo 0
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_GetCachedRiesgoFresh_W1_S6_Adversarial_Concurrent = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' =============================================================================
' ATOM: W1-S7 DB injection verified with isolated temp .accdb
'
' Per access-vba-tdd §5.5: to verify that db injected at
' GetCachedRiesgoFresh propagates to Constructor.getRiesgo (and is not
' replaced by an internal getdb() call), a test must use a separate
' temp .accdb and assert:
'   (a) the write/read goes to the temp
'   (b) the sandbox is untouched
'
' This is the canonical test for db parameter propagation per §5.4 + §5.5.
' =============================================================================
Public Function Test_GetCachedRiesgoFresh_W1_S7_DbInjection_TempAislada() As String
    Dim logs(0 To 11) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_GetCachedRiesgoFresh_W1_S7_DbInjection_TempAislada = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    ' --- Setup: temp .accdb isolated from sandbox ---
    Dim m_TempPath As String
    m_TempPath = Environ$("TEMP") & "\GR_Test_W1S7.accdb"
    On Error Resume Next
    Kill m_TempPath
    On Error GoTo EH

    Dim m_TempDb As DAO.Database
    Set m_TempDb = DBEngine.Workspaces(0).CreateDatabase(m_TempPath, dbLangGeneral, dbVersion120)

    ' --- Create minimum schema for TbRiesgos in temp ---
    m_TempDb.Execute "CREATE TABLE TbRiesgos (" & _
        "IDRiesgo LONG, " & _
        "IDEdicion LONG, " & _
        "CodigoUnico TEXT(50), " & _
        "CodigoRiesgo TEXT(50), " & _
        "Estado TEXT(50))", dbFailOnError

    Const TEST_ID_RIESGO As Long = 900007
    Const TEST_ID_EDICION As Long = 900073

    ' --- Idempotency: clean sandbox range + temp will be recreated ---
    Dim m_SandboxDb As DAO.Database
    Dim m_SandboxErr As String
    Set m_SandboxDb = Test_Fixtures.GetTestDb(m_SandboxErr)
    If m_SandboxDb Is Nothing Then
        m_TempDb.Close: Set m_TempDb = Nothing: Kill m_TempPath
        Test_GetCachedRiesgoFresh_W1_S7_DbInjection_TempAislada = BuildFail("GetTestDb Nothing: " & m_SandboxErr, logs)
        Exit Function
    End If
    m_SandboxDb.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & TEST_ID_RIESGO, dbFailOnError

    ' --- Insert into temp (NOT sandbox) ---
    m_TempDb.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, Estado) " & _
        "VALUES (" & TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", 'PC-W1S7', 'R-W1S7', 'Activo')", dbFailOnError

    logs(logIdx) = "1. Arrange: temp .accdb created at " & m_TempPath
    logIdx = logIdx + 1
    logs(logIdx) = "2. Inserted riesgo IDRiesgo=" & TEST_ID_RIESGO & " in temp (NOT sandbox)"
    logIdx = logIdx + 1

    ' --- Act: call GetCachedRiesgoFresh with db=temp ---
    logs(logIdx) = "3. Act: GetCachedRiesgoFresh(db=temp)"
    logIdx = logIdx + 1
    Dim p_Error As String
    Dim result As riesgo
    Set result = GetCachedRiesgoFresh(CStr(TEST_ID_RIESGO), p_Error, m_TempDb)

    ' --- Assert 1: helper respected db=temp (cache miss reads from temp) ---
    If result Is Nothing Then
        m_TempDb.Close: Set m_TempDb = Nothing: Kill m_TempPath
        Test_GetCachedRiesgoFresh_W1_S7_DbInjection_TempAislada = BuildFail( _
            "result is Nothing — db parameter NOT respected (helper used getdb() internally)", logs)
        Exit Function
    End If
    If CStr(result.IDRiesgo) <> CStr(TEST_ID_RIESGO) Then
        m_TempDb.Close: Set m_TempDb = Nothing: Kill m_TempPath
        Test_GetCachedRiesgoFresh_W1_S7_DbInjection_TempAislada = BuildFail( _
            "result.IDRiesgo mismatch: expected " & TEST_ID_RIESGO & ", got " & result.IDRiesgo, logs)
        Exit Function
    End If
    logs(logIdx) = "4. Assert 1 PASS: result is from temp (db=temp respected)"
    logIdx = logIdx + 1

    ' --- Assert 2: sandbox is untouched (countRows in sandbox = 0) ---
    Dim m_SandboxCount As Long
    Dim rs As DAO.Recordset
    Set rs = m_SandboxDb.OpenRecordset("SELECT COUNT(*) AS c FROM TbRiesgos WHERE IDRiesgo=" & TEST_ID_RIESGO)
    m_SandboxCount = Nz(rs!c, 0)
    rs.Close
    Set rs = Nothing
    If m_SandboxCount <> 0 Then
        m_TempDb.Close: Set m_TempDb = Nothing: Kill m_TempPath
        Test_GetCachedRiesgoFresh_W1_S7_DbInjection_TempAislada = BuildFail( _
            "sandbox has " & m_SandboxCount & " rows for IDRiesgo=" & TEST_ID_RIESGO & " — helper leaked write to sandbox", logs)
        Exit Function
    End If
    logs(logIdx) = "5. Assert 2 PASS: sandbox untouched (countRows=0 for test ID)"
    logIdx = logIdx + 1

    ' --- Assert 3: countRows in temp = 1 (read worked) ---
    Dim m_TempCount As Long
    Set rs = m_TempDb.OpenRecordset("SELECT COUNT(*) AS c FROM TbRiesgos WHERE IDRiesgo=" & TEST_ID_RIESGO)
    m_TempCount = Nz(rs!c, 0)
    rs.Close
    Set rs = Nothing
    If m_TempCount <> 1 Then
        m_TempDb.Close: Set m_TempDb = Nothing: Kill m_TempPath
        Test_GetCachedRiesgoFresh_W1_S7_DbInjection_TempAislada = BuildFail( _
            "temp count=" & m_TempCount & " (expected 1)", logs)
        Exit Function
    End If
    logs(logIdx) = "6. Assert 3 PASS: temp has 1 row (read worked via db=temp)"
    logIdx = logIdx + 1

    logs(logIdx) = "7. VERIFIED: db parameter propagates from GetCachedRiesgoFresh to Constructor.getRiesgo"
    logIdx = logIdx + 1

    ' --- Cleanup ---
    m_TempDb.Close
    Set m_TempDb = Nothing
    Kill m_TempPath

    Test_GetCachedRiesgoFresh_W1_S7_DbInjection_TempAislada = BuildOk("db_injection_verified", logs)
    Exit Function
EH:
    On Error Resume Next
    If Not m_TempDb Is Nothing Then m_TempDb.Close: Set m_TempDb = Nothing
    If Len(m_TempPath) > 0 Then Kill m_TempPath
    On Error GoTo 0
    Dim m_EHMsg As String
    m_EHMsg = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_GetCachedRiesgoFresh_W1_S7_DbInjection_TempAislada = BuildFail(m_EHMsg, logs)
End Function

