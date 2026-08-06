Attribute VB_Name = "Test_RiesgoDetalleRefrescoHelper"
' ============================================================
' Test_RiesgoDetalleRefrescoHelper — RED atoms for
'   modRiesgoDetalleRefrescoHelper.DetalleRefresco_RefrescarManual
'
' Skill: access-vba-tdd v2.5 + access-vba-e2e-methodology
'
' Scope: 1 helper -> 5 scenario classes
'   Happy      > existing riesgo with valid ID; both Object refs populated
'   Sad        > empty ID; non-existent ID
'   Edge       > whitespace-padded ID (no Trim, no such row)
'   Adversarial> call twice in sequence (idempotente)
'
' Helper signature (canonica):
'   Public Function DetalleRefresco_RefrescarManual( _
'       ByVal p_IDRiesgo As String, _
'       Optional ByRef p_ObjRiesgoActivo As Object, _
'       Optional ByRef p_ObjRiesgoAlInicio As Object, _
'       Optional ByRef db As DAO.Database = Nothing, _
'       Optional ByRef p_Error As String) As String
'
' Notes:
'   - The bridge to Form_FormRiesgosGestion is guarded by
'     `If FormularioAbierto(...)`; in tests the form is never open,
'     so the bridge call is skipped and the helper just does cache +
'     base reload.
'   - p_Error convention (Telefonica D&S): empty string on success;
'     populated on real errors (no row found, DAO failure, etc.).
' ============================================================
Option Compare Database
Option Explicit

' ============================================================
' ATOM: Happy — existing riesgo with valid ID; both Object refs
'            populated; returns the ID
'
' Setup:
'   m_ObjUsuarioConectado is whatever ForceLocalBackend configured
'   Seed: 1 riesgo row with ID=900001 in sandbox via Test_Fixtures
' Act: DetalleRefresco_RefrescarManual("900001", objActivo, objAlInicio, p_Error)
' Expect: ok=true, value="900001", objActivo<>Nothing, objAlInicio<>Nothing,
'          p_Error=""
' Teardown: DELETE rows WHERE IDRiesgo BETWEEN 900000 AND 999999
' ============================================================
Public Function Test_DetalleRefresco_RefrescarManual_Happy() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend"
    logIdx = logIdx + 1

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_DetalleRefresco_RefrescarManual_Happy = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_DetalleRefresco_RefrescarManual_Happy = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Const TEST_ID_EDICION As Long = 900050
    Const TEST_ID_PROYECTO As Long = 900060
    Const TEST_ID_RIESGO As Long = 900001

    ' Clean any residual from prior runs (FK order: child -> parent)
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion BETWEEN 900050 AND 900059", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto BETWEEN 900060 AND 900069", dbFailOnError

    ' Seed FK chain: Proyecto -> Edicion -> Riesgo
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-DR-PROJ', 'Test Detalle Refresco project')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) " & _
               "VALUES (" & TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", " & _
               "'PC-DR-900001', 'R-DR-900001')", dbFailOnError

    logs(logIdx) = "2. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO
    logIdx = logIdx + 1

    ' -- Act -------------------------------------------------
    logs(logIdx) = "3. Act: DetalleRefresco_RefrescarManual(" & TEST_ID_RIESGO & ")"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim objActivo As Object
    Dim objAlInicio As Object
    Dim result As String
    result = modRiesgoDetalleRefrescoHelper.DetalleRefresco_RefrescarManual( _
        CStr(TEST_ID_RIESGO), objActivo, objAlInicio, db, p_Error)

    ' -- Assert ----------------------------------------------
    If p_Error <> "" Then
        logs(logIdx) = "4. Assert FAIL: p_Error populated=" & p_Error
        logIdx = logIdx + 1
        Test_DetalleRefresco_RefrescarManual_Happy = BuildFail( _
            "Unexpected error: " & p_Error, logs)
        GoTo Teardown
    End If

    If result <> CStr(TEST_ID_RIESGO) Then
        logs(logIdx) = "4. Assert FAIL: expected " & TEST_ID_RIESGO & ", got '" & result & "'"
        logIdx = logIdx + 1
        Test_DetalleRefresco_RefrescarManual_Happy = BuildFail( _
            "Helper did not return the IDRiesgo", logs)
        GoTo Teardown
    End If

    If objActivo Is Nothing Then
        logs(logIdx) = "4. Assert FAIL: p_ObjRiesgoActivo is Nothing"
        logIdx = logIdx + 1
        Test_DetalleRefresco_RefrescarManual_Happy = BuildFail( _
            "Helper did not populate p_ObjRiesgoActivo", logs)
        GoTo Teardown
    End If

    If objAlInicio Is Nothing Then
        logs(logIdx) = "4. Assert FAIL: p_ObjRiesgoAlInicio is Nothing"
        logIdx = logIdx + 1
        Test_DetalleRefresco_RefrescarManual_Happy = BuildFail( _
            "Helper did not populate p_ObjRiesgoAlInicio", logs)
        GoTo Teardown
    End If

    ' --- Critical assertion: activo and alInicio must be DIFFERENT
    '     object instances. If they are the same instance, the diff
    '     baseline (HasChanges) always returns False and breaks the
    '     save flow. The helper enforces this via "double consume +
    '     invalidation" pattern (see modRiesgoDetalleRefrescoHelper
    '     header docs).
    If objActivo Is objAlInicio Then
        logs(logIdx) = "4. Assert FAIL: objActivo IS objAlInicio (same instance) — diff baseline would be broken"
        logIdx = logIdx + 1
        Test_DetalleRefresco_RefrescarManual_Happy = BuildFail( _
            "p_ObjRiesgoActivo and p_ObjRiesgoAlInicio must be distinct instances", logs)
        GoTo Teardown
    End If

    logs(logIdx) = "4. Assert PASS: returned " & result & " with both Object refs populated AND distinct"
    logIdx = logIdx + 1
    Test_DetalleRefresco_RefrescarManual_Happy = BuildOk(result, logs)
    GoTo Teardown

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_DetalleRefresco_RefrescarManual_Happy = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999", dbFailOnError
    On Error GoTo 0
End Function

' ============================================================
' ATOM: Sad (a) — empty ID; returns ""; p_Error populated
'
' Setup: no seed rows
' Act: DetalleRefresco_RefrescarManual("", ...)
' Expect: ok=true, value="", p_Error contains "vacio"
' ============================================================
Public Function Test_DetalleRefresco_RefrescarManual_Sad_EmptyID() As String
    Dim logs(0 To 7) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_DetalleRefresco_RefrescarManual_Sad_EmptyID = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    logs(logIdx) = "1. Arrange: ForceLocalBackend OK; no seed rows"
    logIdx = logIdx + 1

    logs(logIdx) = "2. Act: DetalleRefresco_RefrescarManual("""")"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim objActivo As Object
    Dim objAlInicio As Object
    Dim result As String
    result = modRiesgoDetalleRefrescoHelper.DetalleRefresco_RefrescarManual( _
        "", objActivo, objAlInicio, , p_Error)

    If result <> "" Then
        logs(logIdx) = "3. Assert FAIL: expected empty string, got '" & result & "'"
        logIdx = logIdx + 1
        Test_DetalleRefresco_RefrescarManual_Sad_EmptyID = BuildFail( _
            "Empty ID must return empty string", logs)
        Exit Function
    End If
    If p_Error = "" Then
        logs(logIdx) = "3. Assert FAIL: expected p_Error populated, got empty"
        logIdx = logIdx + 1
        Test_DetalleRefresco_RefrescarManual_Sad_EmptyID = BuildFail( _
            "Empty ID must populate p_Error", logs)
        Exit Function
    End If
    If InStr(1, p_Error, "vacio", vbTextCompare) = 0 Then
        logs(logIdx) = "3. Assert FAIL: p_Error does not mention 'vacio': " & p_Error
        logIdx = logIdx + 1
        Test_DetalleRefresco_RefrescarManual_Sad_EmptyID = BuildFail( _
            "p_Error should describe empty input", logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: empty ID rejected with p_Error=" & p_Error
    logIdx = logIdx + 1
    Test_DetalleRefresco_RefrescarManual_Sad_EmptyID = BuildOk("rejected_empty_id", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_DetalleRefresco_RefrescarManual_Sad_EmptyID = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM: Sad (b) — non-existent ID; returns ""; p_Error populated
'
' Setup: no seed rows; ensure no residual rows with ID=900099
' Act: DetalleRefresco_RefrescarManual("900099", ...)
' Expect: ok=true, value="", p_Error contains "No se pudo recargar"
' ============================================================
Public Function Test_DetalleRefresco_RefrescarManual_Sad_NonExistent() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_DetalleRefresco_RefrescarManual_Sad_NonExistent = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_DetalleRefresco_RefrescarManual_Sad_NonExistent = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999", dbFailOnError

    logs(logIdx) = "1. Arrange: ForceLocalBackend OK; no seed rows; ID=900099 missing"
    logIdx = logIdx + 1

    Const TEST_ID_MISSING As Long = 900099
    logs(logIdx) = "2. Act: DetalleRefresco_RefrescarManual(" & TEST_ID_MISSING & ")"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim objActivo As Object
    Dim objAlInicio As Object
    Dim result As String
    result = modRiesgoDetalleRefrescoHelper.DetalleRefresco_RefrescarManual( _
        CStr(TEST_ID_MISSING), objActivo, objAlInicio, db, p_Error)

    If result <> "" Then
        logs(logIdx) = "3. Assert FAIL: expected empty string, got '" & result & "'"
        logIdx = logIdx + 1
        Test_DetalleRefresco_RefrescarManual_Sad_NonExistent = BuildFail( _
            "Non-existent ID must return empty string", logs)
        Exit Function
    End If
    If p_Error = "" Then
        logs(logIdx) = "3. Assert FAIL: expected p_Error populated, got empty"
        logIdx = logIdx + 1
        Test_DetalleRefresco_RefrescarManual_Sad_NonExistent = BuildFail( _
            "Non-existent ID must populate p_Error", logs)
        Exit Function
    End If
    If InStr(1, p_Error, "No se pudo recargar", vbTextCompare) = 0 Then
        logs(logIdx) = "3. Assert FAIL: p_Error does not mention reload failure: " & p_Error
        logIdx = logIdx + 1
        Test_DetalleRefresco_RefrescarManual_Sad_NonExistent = BuildFail( _
            "p_Error should describe reload failure", logs)
        Exit Function
    End If

    logs(logIdx) = "3. Assert PASS: non-existent ID rejected with p_Error=" & p_Error
    logIdx = logIdx + 1
    Test_DetalleRefresco_RefrescarManual_Sad_NonExistent = BuildOk("rejected_missing_id", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_DetalleRefresco_RefrescarManual_Sad_NonExistent = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM: Edge — whitespace-padded ID
'
' Setup: seed riesgo with ID=900002 (bare)
' Act: DetalleRefresco_RefrescarManual("  900002  ", ...)
' Expect: Access SQL matches padded ID against bare ID (lenient
'   type coercion). Helper returns bare ID; both refs populated.
'   NOTE: original bridge also passed padded IDs through and got
'   the same result. This atom documents the actual behavior, not
'   a stricter "no Trim" contract the helper does not enforce.
' ============================================================
Public Function Test_DetalleRefresco_RefrescarManual_Edge_Whitespace() As String
    Dim logs(0 To 12) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_DetalleRefresco_RefrescarManual_Edge_Whitespace = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_DetalleRefresco_RefrescarManual_Edge_Whitespace = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Const TEST_ID_EDICION As Long = 900051
    Const TEST_ID_PROYECTO As Long = 900061
    Const TEST_ID_RIESGO As Long = 900002

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion BETWEEN 900050 AND 900059", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto BETWEEN 900060 AND 900069", dbFailOnError

    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-DR-EDGE', 'Edge project')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) " & _
               "VALUES (" & TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", " & _
               "'PC-DR-EDGE', 'R-DR-EDGE')", dbFailOnError

    logs(logIdx) = "1. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO & " (bare)"
    logIdx = logIdx + 1

    Dim paddedId As String
    paddedId = "  " & TEST_ID_RIESGO & "  "
    logs(logIdx) = "2. Act: DetalleRefresco_RefrescarManual(""" & paddedId & """)"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim objActivo As Object
    Dim objAlInicio As Object
    Dim result As String
    result = modRiesgoDetalleRefrescoHelper.DetalleRefresco_RefrescarManual( _
        paddedId, objActivo, objAlInicio, db, p_Error)

    ' --- Access SQL lenient match: padded ID resolves to bare ID.
    If result <> CStr(TEST_ID_RIESGO) Then
        logs(logIdx) = "3. Assert FAIL: expected " & TEST_ID_RIESGO & ", got '" & result & "'"
        logIdx = logIdx + 1
        Test_DetalleRefresco_RefrescarManual_Edge_Whitespace = BuildFail( _
            "Padded ID should resolve to bare ID via Access SQL leniency", logs)
        GoTo Teardown
    End If
    If p_Error <> "" Then
        logs(logIdx) = "3. Assert FAIL: unexpected p_Error=" & p_Error
        logIdx = logIdx + 1
        Test_DetalleRefresco_RefrescarManual_Edge_Whitespace = BuildFail( _
            "Padded ID should not produce error", logs)
        GoTo Teardown
    End If
    If objActivo Is Nothing Or objAlInicio Is Nothing Then
        logs(logIdx) = "3. Assert FAIL: one or both Object refs are Nothing"
        logIdx = logIdx + 1
        Test_DetalleRefresco_RefrescarManual_Edge_Whitespace = BuildFail( _
            "Both Object refs must be populated", logs)
        GoTo Teardown
    End If

    logs(logIdx) = "3. Assert PASS: padded ID resolved to bare ID " & result & "; both refs populated"
    logIdx = logIdx + 1
    Test_DetalleRefresco_RefrescarManual_Edge_Whitespace = BuildOk("padded_id_resolves_via_sql_leniency", logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999", dbFailOnError
    On Error GoTo 0
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_DetalleRefresco_RefrescarManual_Edge_Whitespace = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' ATOM: Adversarial — call twice in sequence (idempotente)
'
' Setup: seed riesgo with ID=900003
' Act: call helper twice with same ID
' Expect: both calls return same ID; both Object refs populated each time
' ============================================================
Public Function Test_DetalleRefresco_RefrescarManual_Adversarial_DobleLlamada() As String
    Dim logs(0 To 12) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_DetalleRefresco_RefrescarManual_Adversarial_DobleLlamada = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_DetalleRefresco_RefrescarManual_Adversarial_DobleLlamada = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Const TEST_ID_EDICION As Long = 900052
    Const TEST_ID_PROYECTO As Long = 900062
    Const TEST_ID_RIESGO As Long = 900003

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion BETWEEN 900050 AND 900059", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto BETWEEN 900060 AND 900069", dbFailOnError

    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO & ", 'TEST-DR-ADV', 'Adversarial project')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) " & _
               "VALUES (" & TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", " & _
               "'PC-DR-ADV', 'R-DR-ADV')", dbFailOnError

    logs(logIdx) = "1. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO
    logIdx = logIdx + 1

    ' -- First call ------------------------------------------
    logs(logIdx) = "2a. Act: first call DetalleRefresco_RefrescarManual(" & TEST_ID_RIESGO & ")"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim objActivo1 As Object
    Dim objAlInicio1 As Object
    Dim result1 As String
    result1 = modRiesgoDetalleRefrescoHelper.DetalleRefresco_RefrescarManual( _
        CStr(TEST_ID_RIESGO), objActivo1, objAlInicio1, db, p_Error)

    If result1 <> CStr(TEST_ID_RIESGO) Or p_Error <> "" _
       Or objActivo1 Is Nothing Or objAlInicio1 Is Nothing Then
        logs(logIdx) = "3a. Assert FAIL: first call failed (result='" & result1 & "', err='" & p_Error & "')"
        logIdx = logIdx + 1
        Test_DetalleRefresco_RefrescarManual_Adversarial_DobleLlamada = BuildFail( _
            "First call must succeed", logs)
        GoTo Teardown
    End If
    logs(logIdx) = "3a. Assert PASS: first call returned " & result1
    logIdx = logIdx + 1

    ' -- Second call -----------------------------------------
    logs(logIdx) = "2b. Act: second call DetalleRefresco_RefrescarManual(" & TEST_ID_RIESGO & ")"
    logIdx = logIdx + 1

    Dim objActivo2 As Object
    Dim objAlInicio2 As Object
    p_Error = ""
    Dim result2 As String
    result2 = modRiesgoDetalleRefrescoHelper.DetalleRefresco_RefrescarManual( _
        CStr(TEST_ID_RIESGO), objActivo2, objAlInicio2, db, p_Error)

    If result2 <> CStr(TEST_ID_RIESGO) Or p_Error <> "" _
       Or objActivo2 Is Nothing Or objAlInicio2 Is Nothing Then
        logs(logIdx) = "3b. Assert FAIL: second call failed (result='" & result2 & "', err='" & p_Error & "')"
        logIdx = logIdx + 1
        Test_DetalleRefresco_RefrescarManual_Adversarial_DobleLlamada = BuildFail( _
            "Second call must also succeed (idempotent)", logs)
        GoTo Teardown
    End If
    logs(logIdx) = "3b. Assert PASS: second call returned " & result2 & " (idempotent)"
    logIdx = logIdx + 1

    Test_DetalleRefresco_RefrescarManual_Adversarial_DobleLlamada = BuildOk("idempotent_ok", logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900000 AND 999999", dbFailOnError
    On Error GoTo 0
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.Description
    logIdx = logIdx + 1
    Test_DetalleRefresco_RefrescarManual_Adversarial_DobleLlamada = BuildFail( _
        "Unhandled exception: " & Err.Description, logs)
End Function

' ============================================================
' Private helpers (BuildOk/BuildFail wrappers - same pattern as
' other test modules in this project)
' ============================================================
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function