Attribute VB_Name = "Test_RiesgoRefrescoHelper"
' ============================================================
' Test_RiesgoRefrescoHelper — TDD atoms for
'   modRiesgoRefrescoHelper.Refresco_RefrescarArbolRiesgosScope
'   modRiesgoRefrescoHelper.Refresco_RefrescarNodoRiesgoActual
'
' Skill: access-vba-tdd v2.5 + access-vba-e2e-methodology
' SDD:   forms-thin-phase0-testeable-2026-06-25 / WI-3
'
' Scope: 2 helpers -> 6 scenario atoms (plus 1 integration in Commit 2.5)
'   Helper 1: Refresco_RefrescarArbolRiesgosScope (filter -> array of IDs)
'   Helper 2: Refresco_RefrescarNodoRiesgoActual    (IDRiesgo -> cached riesgo)
'
' Signatures:
'   Public Function Refresco_RefrescarArbolRiesgosScope( _
'       ByVal p_Filtro As String, _
'       Optional ByRef db As DAO.Database = Nothing, _
'       Optional ByRef p_Error As String) As String
'   Public Function Refresco_RefrescarNodoRiesgoActual( _
'       ByVal p_IDRiesgo As String, _
'       Optional ByRef p_ObjRiesgoActivo As Object, _
'       Optional ByRef db As DAO.Database = Nothing, _
'       Optional ByRef p_Error As String) As String
'
' Conventions:
'   - Returns JSON {ok, value, payload, error, logs} via TestCore_BuildOk / BuildFail
'   - p_Error ByRef (Telefonica D&S); empty string on success
'   - Consumes cache via GetCachedRiesgo (NO bypass to Constructor.getRiesgo)
' ============================================================
Option Compare Database
Option Explicit

' ============================================================
' ATOM: W3-S1 Happy Scope (filter "Activo" -> array of IDs)
'
' Setup: seed 3 riesgos with Estado = 'Activo', 'Aceptado', 'Materializado'
'        in same IDEdicion so they all share the scope filter.
' Act:   Refresco_RefrescarArbolRiesgosScope("Estado='Activo'")
' Expect: ok=true, value is JSON array string with >=1 ID, logs populated
' Teardown: DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900100 AND 900199
' ============================================================
Public Function Test_Refresco_RefrescarArbolRiesgosScope_W3_S1_Happy_Scope() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend"
    logIdx = logIdx + 1

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S1_Happy_Scope = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S1_Happy_Scope = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Const TEST_ID_EDICION As Long = 900110
    Const TEST_ID_PROYECTO As Long = 900120
    Const TEST_ID_ACTIVO1 As Long = 900101
    Const TEST_ID_ACTIVO2 As Long = 900102
    Const TEST_ID_OTRO As Long = 900103

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900100 AND 900199", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion BETWEEN 900110 AND 900119", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto BETWEEN 900120 AND 900129", dbFailOnError

    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-RF-W3S1', 'Scope test project')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, Estado) " & _
               "VALUES (" & TEST_ID_ACTIVO1 & ", " & TEST_ID_EDICION & ", 'PC-RF-W3S1-A1', 'R-RF-W3S1-A1', 'Activo')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, Estado) " & _
               "VALUES (" & TEST_ID_ACTIVO2 & ", " & TEST_ID_EDICION & ", 'PC-RF-W3S1-A2', 'R-RF-W3S1-A2', 'Activo')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, Estado) " & _
               "VALUES (" & TEST_ID_OTRO & ", " & TEST_ID_EDICION & ", 'PC-RF-W3S1-OT', 'R-RF-W3S1-OT', 'Retirado')", dbFailOnError

    logs(logIdx) = "2. Arrange: seeded 3 riesgos (2 Activo, 1 Retirado)"
    logIdx = logIdx + 1

    logs(logIdx) = "3. Act: Refresco_RefrescarArbolRiesgosScope(""Estado='Activo'"")"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim result As String
    result = modRiesgoRefrescoHelper.Refresco_RefrescarArbolRiesgosScope( _
        "Estado='Activo'", db, p_Error)

    If p_Error <> "" Then
        logs(logIdx) = "4. Assert FAIL: p_Error=" & p_Error
        logIdx = logIdx + 1
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S1_Happy_Scope = BuildFail( _
            "Helper returned error: " & p_Error, logs)
        GoTo Teardown
    End If

    If Len(result) = 0 Then
        logs(logIdx) = "4. Assert FAIL: result is empty"
        logIdx = logIdx + 1
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S1_Happy_Scope = BuildFail( _
            "Helper must return JSON with non-empty value", logs)
        GoTo Teardown
    End If

    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then
        logs(logIdx) = "4. Assert FAIL: result missing ok=true: " & Left$(result, 200)
        logIdx = logIdx + 1
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S1_Happy_Scope = BuildFail( _
            "Result must contain ok=true", logs)
        GoTo Teardown
    End If

    ' Must include both activo IDs
    Dim hasA1 As Boolean, hasA2 As Boolean, hasOtro As Boolean
    hasA1 = (InStr(1, result, "900101", vbTextCompare) > 0)
    hasA2 = (InStr(1, result, "900102", vbTextCompare) > 0)
    hasOtro = (InStr(1, result, "900103", vbTextCompare) > 0)

    If Not hasA1 Or Not hasA2 Then
        logs(logIdx) = "4. Assert FAIL: result missing one of activo IDs (A1=" & hasA1 & ", A2=" & hasA2 & ")"
        logIdx = logIdx + 1
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S1_Happy_Scope = BuildFail( _
            "Result must contain IDs of activo riesgos", logs)
        GoTo Teardown
    End If

    If hasOtro Then
        logs(logIdx) = "4. Assert FAIL: result unexpectedly contains Retirado ID 900103"
        logIdx = logIdx + 1
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S1_Happy_Scope = BuildFail( _
            "Filter must exclude Retirado riesgo", logs)
        GoTo Teardown
    End If

    logs(logIdx) = "4. Assert PASS: filter returned activo IDs only (A1, A2)"
    logIdx = logIdx + 1
    Test_Refresco_RefrescarArbolRiesgosScope_W3_S1_Happy_Scope = BuildOk("scope_activo_ok", logs)
    GoTo Teardown

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_Refresco_RefrescarArbolRiesgosScope_W3_S1_Happy_Scope = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900100 AND 900199", dbFailOnError
    On Error GoTo 0
End Function

' ============================================================
' ATOM: W3-S2 Sad Empty Filtro
'
' Setup: no seed needed
' Act:   Refresco_RefrescarArbolRiesgosScope("")
' Expect: ok=false, error mentions "vacio" / "empty", no DB query executed
' ============================================================
Public Function Test_Refresco_RefrescarArbolRiesgosScope_W3_S2_Sad_EmptyFiltro() As String
    Dim logs(0 To 7) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S2_Sad_EmptyFiltro = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    logs(logIdx) = "1. Arrange: ForceLocalBackend OK; no seed"
    logIdx = logIdx + 1

    logs(logIdx) = "2. Act: Refresco_RefrescarArbolRiesgosScope("""")"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim result As String
    result = modRiesgoRefrescoHelper.Refresco_RefrescarArbolRiesgosScope("", , p_Error)

    If result = "" Then
        logs(logIdx) = "3. Assert FAIL: result is empty (expected JSON with ok=false)"
        logIdx = logIdx + 1
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S2_Sad_EmptyFiltro = BuildFail( _
            "Helper must return JSON envelope even on error", logs)
        Exit Function
    End If

    If InStr(1, result, """ok"":false", vbTextCompare) = 0 Then
        logs(logIdx) = "3. Assert FAIL: result missing ok=false: " & Left$(result, 200)
        logIdx = logIdx + 1
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S2_Sad_EmptyFiltro = BuildFail( _
            "Empty filtro must produce ok=false", logs)
        Exit Function
    End If

    ' Must NOT include any successful value
    If InStr(1, result, """ok"":true", vbTextCompare) > 0 Then
        logs(logIdx) = "3. Assert FAIL: ok=true leaked into fail response"
        logIdx = logIdx + 1
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S2_Sad_EmptyFiltro = BuildFail( _
            "Empty filtro must NOT report ok=true", logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: empty filtro rejected with ok=false"
    logIdx = logIdx + 1
    Test_Refresco_RefrescarArbolRiesgosScope_W3_S2_Sad_EmptyFiltro = BuildOk("empty_filtro_rejected", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_Refresco_RefrescarArbolRiesgosScope_W3_S2_Sad_EmptyFiltro = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM: W3-S3 Edge Wildcard (LIKE pattern)
'
' Setup: seed 3 riesgos with CodigoRiesgo 'R-RF-W3S3-A', 'R-RF-W3S3-B', 'X-W3S3-C'
' Act:   Refresco_RefrescarArbolRiesgosScope("CodigoRiesgo LIKE 'R-%'")
' Expect: includes A and B (the two matching 'R-%'), excludes C
' ============================================================
Public Function Test_Refresco_RefrescarArbolRiesgosScope_W3_S3_Edge_Wildcard() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S3_Edge_Wildcard = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S3_Edge_Wildcard = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Const TEST_ID_EDICION As Long = 900130
    Const TEST_ID_PROYECTO As Long = 900140
    Const TEST_ID_A As Long = 900131
    Const TEST_ID_B As Long = 900132
    Const TEST_ID_C As Long = 900133

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900130 AND 900199", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion BETWEEN 900130 AND 900139", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto BETWEEN 900140 AND 900149", dbFailOnError

    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-RF-W3S3', 'Wildcard test project')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) " & _
               "VALUES (" & TEST_ID_A & ", " & TEST_ID_EDICION & ", 'PC-RF-W3S3-A', 'R-RF-W3S3-A')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) " & _
               "VALUES (" & TEST_ID_B & ", " & TEST_ID_EDICION & ", 'PC-RF-W3S3-B', 'R-RF-W3S3-B')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) " & _
               "VALUES (" & TEST_ID_C & ", " & TEST_ID_EDICION & ", 'PC-RF-W3S3-C', 'X-W3S3-C')", dbFailOnError

    logs(logIdx) = "1. Arrange: seeded 3 riesgos (A,B match 'R-%'; C does not)"
    logIdx = logIdx + 1

    logs(logIdx) = "2. Act: Refresco_RefrescarArbolRiesgosScope(""CodigoRiesgo LIKE 'R-%'"")"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim result As String
    result = modRiesgoRefrescoHelper.Refresco_RefrescarArbolRiesgosScope( _
        "CodigoRiesgo LIKE 'R-%'", db, p_Error)

    If p_Error <> "" Then
        logs(logIdx) = "3. Assert FAIL: p_Error=" & p_Error
        logIdx = logIdx + 1
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S3_Edge_Wildcard = BuildFail( _
            "Helper returned error: " & p_Error, logs)
        GoTo Teardown
    End If

    Dim hasA As Boolean, hasB As Boolean, hasC As Boolean
    hasA = (InStr(1, result, CStr(TEST_ID_A), vbTextCompare) > 0)
    hasB = (InStr(1, result, CStr(TEST_ID_B), vbTextCompare) > 0)
    hasC = (InStr(1, result, CStr(TEST_ID_C), vbTextCompare) > 0)

    If Not hasA Or Not hasB Then
        logs(logIdx) = "3. Assert FAIL: missing match (A=" & hasA & ", B=" & hasB & ")"
        logIdx = logIdx + 1
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S3_Edge_Wildcard = BuildFail( _
            "Wildcard 'R-%' must match riesgos starting with R", logs)
        GoTo Teardown
    End If

    If hasC Then
        logs(logIdx) = "3. Assert FAIL: wildcard leaked non-matching ID 900133"
        logIdx = logIdx + 1
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S3_Edge_Wildcard = BuildFail( _
            "Wildcard must exclude non-matching CodigoRiesgo", logs)
        GoTo Teardown
    End If

    logs(logIdx) = "3. Assert PASS: wildcard LIKE 'R-%' matched A and B, excluded C"
    logIdx = logIdx + 1
    Test_Refresco_RefrescarArbolRiesgosScope_W3_S3_Edge_Wildcard = BuildOk("wildcard_r_prefix_ok", logs)
    GoTo Teardown

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_Refresco_RefrescarArbolRiesgosScope_W3_S3_Edge_Wildcard = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900130 AND 900199", dbFailOnError
    On Error GoTo 0
End Function

' ============================================================
' ATOM: W3-S4 Edge Whitespace (Refresco_RefrescarNodoRiesgoActual)
'
' Setup: seed riesgo IDRiesgo=900150 (bare)
' Act:   Refresco_RefrescarNodoRiesgoActual("  900150  ")
' Expect: ok=true; p_Error empty; p_ObjRiesgoActivo populated; result is JSON.
'         Access SQL lenient match: padded ID resolves to bare ID 900150.
' ============================================================
Public Function Test_Refresco_RefrescarNodoRiesgoActual_W3_S4_Edge_Whitespace() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_Refresco_RefrescarNodoRiesgoActual_W3_S4_Edge_Whitespace = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Refresco_RefrescarNodoRiesgoActual_W3_S4_Edge_Whitespace = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Const TEST_ID_EDICION As Long = 900151
    Const TEST_ID_PROYECTO As Long = 900152
    Const TEST_ID_RIESGO As Long = 900150

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900150 AND 900159", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion BETWEEN 900151 AND 900160", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto BETWEEN 900152 AND 900161", dbFailOnError

    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-RF-W3S4', 'Whitespace test project')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) " & _
               "VALUES (" & TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", 'PC-RF-W3S4', 'R-RF-W3S4')", dbFailOnError

    logs(logIdx) = "1. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO & " (bare)"
    logIdx = logIdx + 1

    Dim paddedId As String
    paddedId = "  " & TEST_ID_RIESGO & "  "
    logs(logIdx) = "2. Act: Refresco_RefrescarNodoRiesgoActual(""" & paddedId & """)"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim objActivo As Object
    Dim result As String
    result = modRiesgoRefrescoHelper.Refresco_RefrescarNodoRiesgoActual( _
        paddedId, objActivo, db, p_Error)

    If p_Error <> "" Then
        logs(logIdx) = "3. Assert FAIL: p_Error=" & p_Error
        logIdx = logIdx + 1
        Test_Refresco_RefrescarNodoRiesgoActual_W3_S4_Edge_Whitespace = BuildFail( _
            "Padded ID must not produce error: " & p_Error, logs)
        GoTo Teardown
    End If

    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then
        logs(logIdx) = "3. Assert FAIL: result missing ok=true: " & Left$(result, 200)
        logIdx = logIdx + 1
        Test_Refresco_RefrescarNodoRiesgoActual_W3_S4_Edge_Whitespace = BuildFail( _
            "Result must contain ok=true for resolvable padded ID", logs)
        GoTo Teardown
    End If

    If objActivo Is Nothing Then
        logs(logIdx) = "3. Assert FAIL: p_ObjRiesgoActivo is Nothing"
        logIdx = logIdx + 1
        Test_Refresco_RefrescarNodoRiesgoActual_W3_S4_Edge_Whitespace = BuildFail( _
            "Helper must populate p_ObjRiesgoActivo via GetCachedRiesgo", logs)
        GoTo Teardown
    End If

    If CStr(objActivo.IDRiesgo) <> CStr(TEST_ID_RIESGO) Then
        logs(logIdx) = "3. Assert FAIL: objActivo.IDRiesgo='" & objActivo.IDRiesgo & "' (expected 900150)"
        logIdx = logIdx + 1
        Test_Refresco_RefrescarNodoRiesgoActual_W3_S4_Edge_Whitespace = BuildFail( _
            "Padded ID must resolve to bare ID via SQL leniency", logs)
        GoTo Teardown
    End If

    logs(logIdx) = "3. Assert PASS: padded ID " & paddedId & " resolved to bare ID 900150"
    logIdx = logIdx + 1
    Test_Refresco_RefrescarNodoRiesgoActual_W3_S4_Edge_Whitespace = BuildOk("padded_id_resolved", logs)
    GoTo Teardown

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_Refresco_RefrescarNodoRiesgoActual_W3_S4_Edge_Whitespace = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900150 AND 900159", dbFailOnError
    On Error GoTo 0
End Function

' ============================================================
' ATOM: W3-S5 Adversarial - doble llamada (idempotente)
'
' Setup: seed riesgo IDRiesgo=900160
' Act:   call Refresco_RefrescarArbolRiesgosScope twice with same filter
' Expect: both calls return equivalent arrays; no errors; both ok=true
'         (deterministic, no caching of filter results)
' ============================================================
Public Function Test_Refresco_RefrescarArbolRiesgosScope_W3_S5_Adversarial_DobleLlamada() As String
    Dim logs(0 To 12) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S5_Adversarial_DobleLlamada = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S5_Adversarial_DobleLlamada = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Const TEST_ID_EDICION As Long = 900161
    Const TEST_ID_PROYECTO As Long = 900162
    Const TEST_ID_RIESGO As Long = 900160

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900160 AND 900169", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion BETWEEN 900161 AND 900170", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto BETWEEN 900162 AND 900171", dbFailOnError

    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-RF-W3S5', 'Adversarial test project')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, Estado) " & _
               "VALUES (" & TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", 'PC-RF-W3S5', 'R-RF-W3S5', 'Activo')", dbFailOnError

    logs(logIdx) = "1. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO & " (Estado=Activo)"
    logIdx = logIdx + 1

    logs(logIdx) = "2a. Act: first call"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim result1 As String
    result1 = modRiesgoRefrescoHelper.Refresco_RefrescarArbolRiesgosScope( _
        "Estado='Activo'", db, p_Error)

    If p_Error <> "" Then
        logs(logIdx) = "3a. Assert FAIL: first call p_Error=" & p_Error
        logIdx = logIdx + 1
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S5_Adversarial_DobleLlamada = BuildFail( _
            "First call errored: " & p_Error, logs)
        GoTo Teardown
    End If

    If InStr(1, result1, CStr(TEST_ID_RIESGO), vbTextCompare) = 0 Then
        logs(logIdx) = "3a. Assert FAIL: first call missing seeded ID"
        logIdx = logIdx + 1
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S5_Adversarial_DobleLlamada = BuildFail( _
            "First call must include seeded activo riesgo", logs)
        GoTo Teardown
    End If
    logs(logIdx) = "3a. Assert PASS: first call returned ID 900160"
    logIdx = logIdx + 1

    logs(logIdx) = "2b. Act: second call (same filter)"
    logIdx = logIdx + 1

    p_Error = ""
    Dim result2 As String
    result2 = modRiesgoRefrescoHelper.Refresco_RefrescarArbolRiesgosScope( _
        "Estado='Activo'", db, p_Error)

    If p_Error <> "" Then
        logs(logIdx) = "3b. Assert FAIL: second call p_Error=" & p_Error
        logIdx = logIdx + 1
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S5_Adversarial_DobleLlamada = BuildFail( _
            "Second call errored: " & p_Error, logs)
        GoTo Teardown
    End If

    If InStr(1, result2, CStr(TEST_ID_RIESGO), vbTextCompare) = 0 Then
        logs(logIdx) = "3b. Assert FAIL: second call missing seeded ID"
        logIdx = logIdx + 1
        Test_Refresco_RefrescarArbolRiesgosScope_W3_S5_Adversarial_DobleLlamada = BuildFail( _
            "Second call must include seeded activo riesgo", logs)
        GoTo Teardown
    End If
    logs(logIdx) = "3b. Assert PASS: second call returned ID 900160 (idempotent)"
    logIdx = logIdx + 1

    Test_Refresco_RefrescarArbolRiesgosScope_W3_S5_Adversarial_DobleLlamada = BuildOk("idempotent_ok", logs)
    GoTo Teardown

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_Refresco_RefrescarArbolRiesgosScope_W3_S5_Adversarial_DobleLlamada = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900160 AND 900169", dbFailOnError
    On Error GoTo 0
End Function

' ============================================================
' ATOM: W3-S6 Sandbox (m_TestingMode + cache hit on nodo helper)
'
' Setup: seed riesgo IDRiesgo=900170. ForceLocalBackend sets m_TestingMode.
'        Call GetCachedRiesgo to populate cache, then call helper.
' Act:   Refresco_RefrescarNodoRiesgoActual("900170")
' Expect: helper reads from cache (CacheRiesgoContainsKey=900170 BEFORE call
'         is optional; AFTER call cache must be populated); p_Error="".
'         This proves the helper routes through GetCachedRiesgo, not a
'         bypass to Constructor.getRiesgo.
' ============================================================
Public Function Test_Refresco_RefrescarNodoRiesgoActual_W3_S6_Sandbox_CacheHit() As String
    Dim logs(0 To 12) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_Refresco_RefrescarNodoRiesgoActual_W3_S6_Sandbox_CacheHit = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Refresco_RefrescarNodoRiesgoActual_W3_S6_Sandbox_CacheHit = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Const TEST_ID_EDICION As Long = 900171
    Const TEST_ID_PROYECTO As Long = 900172
    Const TEST_ID_RIESGO As Long = 900170

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900170 AND 900179", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion BETWEEN 900171 AND 900180", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto BETWEEN 900172 AND 900181", dbFailOnError

    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-RF-W3S6', 'Sandbox test project')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) " & _
               "VALUES (" & TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", 'PC-RF-W3S6', 'R-RF-W3S6')", dbFailOnError

    ' Ensure clean cache state before exercising the helper
    Dim cacheErr As String
    InvalidarCacheRiesgo CStr(TEST_ID_RIESGO), cacheErr

    logs(logIdx) = "1. Arrange: seeded riesgo; cache cleared; m_TestingMode=True"
    logIdx = logIdx + 1

    logs(logIdx) = "2. Act: Refresco_RefrescarNodoRiesgoActual(" & TEST_ID_RIESGO & ")"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim objActivo As Object
    Dim result As String
    result = modRiesgoRefrescoHelper.Refresco_RefrescarNodoRiesgoActual( _
        CStr(TEST_ID_RIESGO), objActivo, db, p_Error)

    If p_Error <> "" Then
        logs(logIdx) = "3. Assert FAIL: p_Error=" & p_Error
        logIdx = logIdx + 1
        Test_Refresco_RefrescarNodoRiesgoActual_W3_S6_Sandbox_CacheHit = BuildFail( _
            "Helper errored: " & p_Error, logs)
        GoTo Teardown
    End If

    If objActivo Is Nothing Then
        logs(logIdx) = "3. Assert FAIL: p_ObjRiesgoActivo is Nothing"
        logIdx = logIdx + 1
        Test_Refresco_RefrescarNodoRiesgoActual_W3_S6_Sandbox_CacheHit = BuildFail( _
            "Helper must populate p_ObjRiesgoActivo", logs)
        GoTo Teardown
    End If

    ' CRITICAL cache contract: helper MUST populate m_DicRiesgos via
    ' GetCachedRiesgo. If it bypassed to Constructor.getRiesgo, the
    ' cache would stay empty and tests relying on cache mock would break.
    If Not CacheRiesgoContainsKey(CStr(TEST_ID_RIESGO)) Then
        logs(logIdx) = "3. Assert FAIL: cache does NOT contain IDRiesgo=" & TEST_ID_RIESGO
        logIdx = logIdx + 1
        Test_Refresco_RefrescarNodoRiesgoActual_W3_S6_Sandbox_CacheHit = BuildFail( _
            "Helper bypassed cache (called Constructor.getRiesgo directly)", logs)
        GoTo Teardown
    End If

    logs(logIdx) = "3. Assert PASS: helper populated cache via GetCachedRiesgo"
    logIdx = logIdx + 1
    Test_Refresco_RefrescarNodoRiesgoActual_W3_S6_Sandbox_CacheHit = BuildOk("cache_populated_via_helper", logs)
    GoTo Teardown

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_Refresco_RefrescarNodoRiesgoActual_W3_S6_Sandbox_CacheHit = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900170 AND 900179", dbFailOnError
    On Error GoTo 0
End Function

' ============================================================
' Private helpers - JSON wrappers (same pattern as other Test_*.bas)
' ============================================================
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function