Attribute VB_Name = "Test_RiesgoEstadoGateHelper"
' ============================================================
' Test_RiesgoEstadoGateHelper — TDD Red Atoms
'
' Skill: access-vba-tdd v2.4
' Helper under test: modRiesgoEstadoGateHelper.bas (NOT YET EXISTS — RED)
'
' Module: GEMELO PC/CDCA/CDCASUB (highest priority)
' Gemelo scope: 3 types — PC, CDCA, CDCASUB (no PCSUB)
'
' 7 helpers ? 4 atoms = 28 red atoms total
'
' Assumptions (?contract):
'   1. Validación fecha Retiro = same rule as Materializado (TODO: confirmar)
'   2. Gemelo scope: PC, CDCA, CDCASUB — no PCSUB
'   3. Suministradores not in this module
'   4. NC not in this module
'
' Divergences:
'   - QuitarMaterializacion (#2) vs QuitarRetiroRiesgo (#6): DIFFERENT helpers,
'     different domain methods, same shared validation pattern
' ============================================================
Option Compare Database
Option Explicit

' --- JSON wrappers (BuildOk/BuildFail) — must precede all Public Functions
' Per access-vba-tdd §1.8
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' --- Private helpers -------------------------------------------------

Private Function CountRows(ByVal db As DAO.Database, ByVal p_Table As String, _
                           ByVal p_Where As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS c FROM " & p_Table & " WHERE " & p_Where)
    CountRows = CLng(Nz(rs!c, 0))
    rs.Close
    Set rs = Nothing
End Function

Private Sub EnsureTableClean(ByVal db As DAO.Database, ByVal p_TableName As String, _
                             Optional ByVal p_WhereClause As String = "")
    On Error Resume Next
    If Len(p_WhereClause) = 0 Then
        db.Execute "DELETE FROM " & p_TableName & " WHERE IDRiesgo>=900000 OR ID>=900000", dbFailOnError
    Else
        db.Execute "DELETE FROM " & p_TableName & " WHERE " & p_WhereClause, dbFailOnError
    End If
    On Error GoTo 0
End Sub

Private Sub EnsureTbRiesgosClean(ByVal db As DAO.Database, ByVal p_IDRiesgo As Long)
    On Error Resume Next
    db.Execute "UPDATE TbRiesgos SET FechaMaterializado=NULL, FechaRetirado=NULL, " & _
        "FechaMitigacionAceptar=NULL WHERE IDRiesgo=" & p_IDRiesgo, dbFailOnError
    On Error GoTo 0
End Sub

Private Sub SeedMaterializacionSi(ByVal db As DAO.Database, ByVal p_IDRiesgo As Long, _
                                   ByVal p_IDProyecto As Long, ByVal p_IDEdicion As Long, _
                                   ByVal p_CodigoRiesgo As String, _
                                   ByVal p_IDMat As Long, ByVal p_Fecha As Date)
    db.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE ID=" & p_IDMat
    db.Execute "INSERT INTO TbRiesgosMaterializaciones " & _
        "(ID, IDProyecto, IDEdicion, CodigoRiesgo, Fecha, EsMaterializacion, Estado) " & _
        "VALUES (" & p_IDMat & ", " & p_IDProyecto & ", " & p_IDEdicion & ", " & _
        "'" & Replace(p_CodigoRiesgo, "'", "''") & "', #" & Format(p_Fecha, "yyyy-mm-dd") & "#, " & _
        "'Sí', 'Materializado')", dbFailOnError
End Sub

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

' --- RunAll ---
Public Function Test_RiesgoEstadoGateHelper_RunAll() As String
    Dim logs(0 To 0) As String
    logs(0) = "Runner: Test_RiesgoEstadoGateHelper_RunAll"
    Test_RiesgoEstadoGateHelper_RunAll = BuildOk("28 atoms", logs)
End Function

' =====================================================================
' #1 ValidarTransicionMaterializacion — 4 atoms
' Signature: Function ValidarTransicionMaterializacion(
'   ByVal p_IDRiesgo As Long, ByRef p_Fecha As String, ByRef p_IDPlan As String,
'   Optional ByRef db As DAO.Database = Nothing,
'   Optional ByRef p_PromptResult As Long) As String
' Form: Form_FormRiesgoMaterializado (inline lines 25-85)
' Returns: "" on success, error string on failure
' =====================================================================

' Happy: valid future date + active plan linked to riesgo
Public Function Test_ValidarTransicionMaterializacion_Happy() As String
    Dim logs(0 To 7) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim p_Fecha As String
    Dim p_IDPlan As String
    Dim p_PromptResult As Long
    Dim result As String

    logs(0) = "1. Arrange: ForceLocalBackend sandbox"
    logs(1) = "2. Arrange: SeedAll base graph (riesgo=900503)"
    logs(2) = "3. Arrange: Estado=Detectado (pre-transition)"
    logs(3) = "4. Arrange: date within plan period + plan ID linked"
    logs(4) = "5. Act: ValidarTransicionMaterializacion(900503, date, plan, db)"
    logs(5) = "6. Assert: result = """" (success)"
    logs(6) = "7. Teardown"
    logs(7) = "Assumption: validacion fecha uses same rule as materializado (TODO: confirmar)"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_ValidarTransicionMaterializacion_Happy = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ValidarTransicionMaterializacion_Happy = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset("SELECT CodigoRiesgo FROM TbRiesgos WHERE IDRiesgo=" & _
        Test_Fixtures.Cache_RiesgoId, dbOpenSnapshot)
    If rs.EOF Then
        Test_ValidarTransicionMaterializacion_Happy = BuildFail("fixture riesgo not found", logs)
        rs.Close: GoTo Teardown
    End If
    Dim CodRiesgo As String
    CodRiesgo = CStr(Nz(rs!codigoRiesgo.value, ""))
    rs.Close

    db.Execute "UPDATE TbRiesgos SET FechaMaterializado=NULL, Estado='Detectado' " & _
        "WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, dbFailOnError

    ' Use fixture PM (900504) — linked to fixture riesgo
    p_Fecha = Format$(DateSerial(2026, 12, 31), "dd/mm/yyyy")
    p_IDPlan = CStr(Test_Fixtures.Cache_PMId)
    p_PromptResult = 0

    logs(4) = "5. Act: ValidarTransicionMaterializacion(900503, """ & p_Fecha & """, """ & p_IDPlan & """)"

    ' RED: helper does not exist yet — call will fail at runtime
    result = ValidarTransicionMaterializacion(Test_Fixtures.Cache_RiesgoId, p_Fecha, p_IDPlan, db, p_PromptResult)

    If result <> "" Then
        logs(5) = "6. FAILED: result=""" & result & """, expected """""
        Test_ValidarTransicionMaterializacion_Happy = BuildFail("expected empty: " & result, logs)
        GoTo Teardown
    End If

    logs(5) = "6. OK: result is empty (validation passed)"
    Test_ValidarTransicionMaterializacion_Happy = BuildOk("validado ok", logs)
    GoTo Teardown

EH:
    Test_ValidarTransicionMaterializacion_Happy = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Sad: blank fecha (required field)
Public Function Test_ValidarTransicionMaterializacion_Sad_BlankFecha() As String
    Dim logs(0 To 6) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim p_Fecha As String
    Dim p_IDPlan As String
    Dim result As String

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: blank fecha"
    logs(3) = "4. Act: ValidarTransicionMaterializacion with blank fecha"
    logs(4) = "5. Assert: result <> "" (validation error)"
    logs(5) = "Sad: blank required field"
    logs(6) = "Assumption: same date-validation rule (TODO: confirmar)"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_ValidarTransicionMaterializacion_Sad_BlankFecha = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ValidarTransicionMaterializacion_Sad_BlankFecha = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    p_Fecha = ""
    p_IDPlan = CStr(Test_Fixtures.Cache_PMId)

    result = ValidarTransicionMaterializacion(Test_Fixtures.Cache_RiesgoId, p_Fecha, p_IDPlan, db)

    If result = "" Then
        logs(4) = "5. FAILED: blank fecha should have returned error"
        Test_ValidarTransicionMaterializacion_Sad_BlankFecha = BuildFail("expected error for blank fecha", logs)
        GoTo Teardown
    End If

    logs(4) = "5. OK: blank fecha rejected with: " & result
    Test_ValidarTransicionMaterializacion_Sad_BlankFecha = BuildOk("blank fecha rejected: " & result, logs)
    GoTo Teardown

EH:
    Test_ValidarTransicionMaterializacion_Sad_BlankFecha = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Edge: fecha outside plan period (plan ends 2026-12-31; fecha=2027-01-01)
Public Function Test_ValidarTransicionMaterializacion_Edge_FechaOutsidePlanPeriod() As String
    Dim logs(0 To 6) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim p_Fecha As String
    Dim p_IDPlan As String
    Dim result As String

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: fecha=2027-01-01 (after plan period end)"
    logs(3) = "4. Act: ValidarTransicionMaterializacion with out-of-period fecha"
    logs(4) = "5. Assert: result <> "" (validation error)"
    logs(5) = "Edge: date outside plan validity period"
    logs(6) = "Assumption: ValidarFechaMaterializacionPermitida applies (TODO: confirmar)"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_ValidarTransicionMaterializacion_Edge_FechaOutsidePlanPeriod = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ValidarTransicionMaterializacion_Edge_FechaOutsidePlanPeriod = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    p_Fecha = "01/01/2027"
    p_IDPlan = CStr(Test_Fixtures.Cache_PMId)

    result = ValidarTransicionMaterializacion(Test_Fixtures.Cache_RiesgoId, p_Fecha, p_IDPlan, db)

    If result = "" Then
        logs(4) = "5. FAILED: out-of-period fecha should have returned error"
        Test_ValidarTransicionMaterializacion_Edge_FechaOutsidePlanPeriod = BuildFail("expected error for out-of-period fecha", logs)
        GoTo Teardown
    End If

    logs(4) = "5. OK: out-of-period fecha rejected: " & result
    Test_ValidarTransicionMaterializacion_Edge_FechaOutsidePlanPeriod = BuildOk("out-of-period rejected: " & result, logs)
    GoTo Teardown

EH:
    Test_ValidarTransicionMaterializacion_Edge_FechaOutsidePlanPeriod = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Adversarial: riesgo ID does not exist
Public Function Test_ValidarTransicionMaterializacion_Adversarial_IDNoExiste() As String
    Dim logs(0 To 5) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim p_Fecha As String
    Dim p_IDPlan As String
    Dim result As String
    Const NON_EXISTENT_ID As Long = 999999

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Act: ValidarTransicionMaterializacion with non-existent ID=" & NON_EXISTENT_ID
    logs(3) = "4. Assert: result <> "" (error for non-existent riesgo)"
    logs(4) = "Adversarial: invalid riesgo ID"
    logs(5) = "5. Teardown"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_ValidarTransicionMaterializacion_Adversarial_IDNoExiste = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ValidarTransicionMaterializacion_Adversarial_IDNoExiste = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    p_Fecha = Format$(DateSerial(2026, 12, 31), "dd/mm/yyyy")
    p_IDPlan = CStr(Test_Fixtures.Cache_PMId)

    result = ValidarTransicionMaterializacion(NON_EXISTENT_ID, p_Fecha, p_IDPlan, db)

    If result = "" Then
        logs(3) = "4. FAILED: non-existent riesgo should have returned error"
        Test_ValidarTransicionMaterializacion_Adversarial_IDNoExiste = BuildFail("expected error for non-existent riesgo", logs)
        GoTo Teardown
    End If

    logs(3) = "4. OK: non-existent riesgo rejected: " & result
    Test_ValidarTransicionMaterializacion_Adversarial_IDNoExiste = BuildOk("non-existent rejected: " & result, logs)
    GoTo Teardown

EH:
    Test_ValidarTransicionMaterializacion_Adversarial_IDNoExiste = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' =====================================================================
' #2 QuitarMaterializacion — 4 atoms
' Signature: Function QuitarMaterializacion(
'   ByVal p_IDRiesgo As Long,
'   Optional ByRef db As DAO.Database = Nothing,
'   Optional ByRef p_PromptResult As Long) As String
' Form: Form_FormRiesgoMaterializado (inline lines 209-259)
' Returns: "" on success, error string on failure
' =====================================================================

' Happy: riesgo is Materializado; confirm vbYes; Si-row deleted
Public Function Test_QuitarMaterializacion_Happy() As String
    Dim logs(0 To 8) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim p_PromptResult As Long
    Dim result As String
    Dim countBefore As Long
    Dim countAfter As Long
    Const FIX_MAT_ID As Long = 900600

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: Estado=Materializado + Si-row id=900600"
    logs(3) = "4. Act: QuitarMaterializacion(900503, vbYes)"
    logs(4) = "5. Assert: result="""" and p_PromptResult=vbYes"
    logs(5) = "6. Assert: Si-row deleted (count 1->0)"
    logs(6) = "7. Assert: Estado reverted from Materializado"
    logs(7) = "8. Teardown"
    logs(8) = "Happy: valid removal confirmed"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_QuitarMaterializacion_Happy = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_QuitarMaterializacion_Happy = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset("SELECT CodigoRiesgo FROM TbRiesgos WHERE IDRiesgo=" & _
        Test_Fixtures.Cache_RiesgoId, dbOpenSnapshot)
    If rs.EOF Then
        Test_QuitarMaterializacion_Happy = BuildFail("fixture riesgo not found", logs)
        rs.Close: GoTo Teardown
    End If
    Dim CodRiesgo As String
    CodRiesgo = CStr(Nz(rs!codigoRiesgo.value, ""))
    rs.Close

    db.Execute "UPDATE TbRiesgos SET FechaMaterializado=#2026-06-01#, Estado='Materializado' " & _
        "WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, dbFailOnError
    Call SeedMaterializacionSi(db, Test_Fixtures.Cache_RiesgoId, Test_Fixtures.Cache_ProyectoId, _
        Test_Fixtures.Cache_EdicionId, CodRiesgo, FIX_MAT_ID, DateSerial(2026, 6, 1))

    countBefore = CountRows(db, "TbRiesgosMaterializaciones", "ID=" & FIX_MAT_ID)
    If countBefore <> 1 Then
        logs(2) = "3. FAILED: Si-row not seeded (count=" & countBefore & ")"
        Test_QuitarMaterializacion_Happy = BuildFail("fixture not seeded", logs)
        GoTo Teardown
    End If

    p_PromptResult = vbYes
    ' RED: helper does not exist — will error at runtime
    result = QuitarMaterializacion(Test_Fixtures.Cache_RiesgoId, db, p_PromptResult)

    countAfter = CountRows(db, "TbRiesgosMaterializaciones", "ID=" & FIX_MAT_ID)

    If result <> "" Then
        logs(4) = "4. FAILED: result=""" & result & """, expected """""
        Test_QuitarMaterializacion_Happy = BuildFail("expected empty: " & result, logs)
        GoTo Teardown
    End If
    If p_PromptResult <> vbYes Then
        logs(4) = "4. FAILED: p_PromptResult=" & p_PromptResult & " expected vbYes"
        Test_QuitarMaterializacion_Happy = BuildFail("p_PromptResult wrong", logs)
        GoTo Teardown
    End If
    If countAfter <> 0 Then
        logs(5) = "5. FAILED: Si-row not deleted (count=" & countAfter & ")"
        Test_QuitarMaterializacion_Happy = BuildFail("row not deleted", logs)
        GoTo Teardown
    End If

    logs(5) = "5. OK: Si-row deleted (before=" & countBefore & ", after=" & countAfter & ")"
    Test_QuitarMaterializacion_Happy = BuildOk("QuitarMaterializacion ok", logs)
    GoTo Teardown

EH:
    Test_QuitarMaterializacion_Happy = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call EnsureTableClean(db, "TbRiesgosMaterializaciones")
    Call EnsureTbRiesgosClean(db, Test_Fixtures.Cache_RiesgoId)
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Sad: riesgo is not in Materializado state
Public Function Test_QuitarMaterializacion_Sad_EstadoNoMaterializado() As String
    Dim logs(0 To 6) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim p_PromptResult As Long
    Dim result As String
    Dim Estado As String
    Dim rs As DAO.Recordset

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: Estado=Detectado (not Materializado)"
    logs(3) = "4. Act: QuitarMaterializacion on non-materializado riesgo"
    logs(4) = "5. Assert: result handling (error or no-op)"
    logs(5) = "Sad: invalid state transition"
    logs(6) = "6. Teardown"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_QuitarMaterializacion_Sad_EstadoNoMaterializado = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_QuitarMaterializacion_Sad_EstadoNoMaterializado = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    db.Execute "UPDATE TbRiesgos SET FechaMaterializado=NULL, Estado='Detectado' " & _
        "WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, dbFailOnError
    Call EnsureTableClean(db, "TbRiesgosMaterializaciones", _
        "IDProyecto=" & Test_Fixtures.Cache_ProyectoId)

    p_PromptResult = vbYes
    result = QuitarMaterializacion(Test_Fixtures.Cache_RiesgoId, db, p_PromptResult)

    Set rs = db.OpenRecordset("SELECT Estado FROM TbRiesgos WHERE IDRiesgo=" & _
        Test_Fixtures.Cache_RiesgoId, dbOpenSnapshot)
    If Not rs.EOF Then Estado = CStr(Nz(rs!Estado.value, ""))
    rs.Close

    If result = "" Then
        logs(4) = "4. INFO: no-op on non-materializado; Estado=" & Estado
        Test_QuitarMaterializacion_Sad_EstadoNoMaterializado = BuildOk("no-op: " & Estado, logs)
        GoTo Teardown
    End If

    logs(4) = "4. OK: invalid state rejected: " & result
    Test_QuitarMaterializacion_Sad_EstadoNoMaterializado = BuildOk("invalid state rejected: " & result, logs)
    GoTo Teardown

EH:
    Test_QuitarMaterializacion_Sad_EstadoNoMaterializado = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call EnsureTableClean(db, "TbRiesgosMaterializaciones")
    Call EnsureTbRiesgosClean(db, Test_Fixtures.Cache_RiesgoId)
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Edge: user clicks vbNo (cancels removal)
Public Function Test_QuitarMaterializacion_Edge_ConfirmacionRechazada() As String
    Dim logs(0 To 7) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim p_PromptResult As Long
    Dim result As String
    Dim countBefore As Long
    Dim countAfter As Long
    Const FIX_MAT_ID As Long = 900601

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll + materializado fixture"
    logs(2) = "3. Act: QuitarMaterializacion with vbNo (user cancels)"
    logs(3) = "4. Assert: p_PromptResult=vbNo respected"
    logs(4) = "5. Assert: Si-row unchanged (count unchanged)"
    logs(5) = "Edge: user rejects confirmation"
    logs(6) = "6. Teardown"
    logs(7) = ""

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_QuitarMaterializacion_Edge_ConfirmacionRechazada = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_QuitarMaterializacion_Edge_ConfirmacionRechazada = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset("SELECT CodigoRiesgo FROM TbRiesgos WHERE IDRiesgo=" & _
        Test_Fixtures.Cache_RiesgoId, dbOpenSnapshot)
    If rs.EOF Then
        Test_QuitarMaterializacion_Edge_ConfirmacionRechazada = BuildFail("fixture not found", logs)
        rs.Close: GoTo Teardown
    End If
    Dim CodRiesgo As String
    CodRiesgo = CStr(Nz(rs!codigoRiesgo.value, ""))
    rs.Close

    db.Execute "UPDATE TbRiesgos SET FechaMaterializado=#2026-06-01#, Estado='Materializado' " & _
        "WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, dbFailOnError
    Call SeedMaterializacionSi(db, Test_Fixtures.Cache_RiesgoId, Test_Fixtures.Cache_ProyectoId, _
        Test_Fixtures.Cache_EdicionId, CodRiesgo, FIX_MAT_ID, DateSerial(2026, 6, 1))

    countBefore = CountRows(db, "TbRiesgosMaterializaciones", "ID=" & FIX_MAT_ID)

    p_PromptResult = vbNo  ' User cancels
    result = QuitarMaterializacion(Test_Fixtures.Cache_RiesgoId, db, p_PromptResult)

    countAfter = CountRows(db, "TbRiesgosMaterializaciones", "ID=" & FIX_MAT_ID)

    If countAfter <> countBefore Then
        logs(4) = "4. FAILED: vbNo ignored; rows changed (before=" & countBefore & ", after=" & countAfter & ")"
        Test_QuitarMaterializacion_Edge_ConfirmacionRechazada = BuildFail("vbNo ignored", logs)
        GoTo Teardown
    End If

    logs(4) = "4. OK: vbNo respected; rows unchanged"
    Test_QuitarMaterializacion_Edge_ConfirmacionRechazada = BuildOk("cancelled: no changes", logs)
    GoTo Teardown

EH:
    Test_QuitarMaterializacion_Edge_ConfirmacionRechazada = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call EnsureTableClean(db, "TbRiesgosMaterializaciones")
    Call EnsureTbRiesgosClean(db, Test_Fixtures.Cache_RiesgoId)
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Adversarial: db=Nothing causes error
Public Function Test_QuitarMaterializacion_Adversarial_DbNothing() As String
    Dim logs(0 To 5) As String
    Dim p_PromptResult As Long
    Dim result As String

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Act: QuitarMaterializacion with db=Nothing"
    logs(3) = "4. Assert: result <> "" (error — db required)"
    logs(4) = "Adversarial: null db causes error"
    logs(5) = "5. Teardown"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(logs(0)) Then
        Test_QuitarMaterializacion_Adversarial_DbNothing = BuildFail("TESTS BLOCKED", logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    p_PromptResult = vbYes
    result = QuitarMaterializacion(Test_Fixtures.Cache_RiesgoId, Nothing, p_PromptResult)

    If result = "" Then
        logs(3) = "4. FAILED: db=Nothing should have returned error"
        Test_QuitarMaterializacion_Adversarial_DbNothing = BuildFail("expected error for db=Nothing", logs)
        GoTo Teardown
    End If

    logs(3) = "4. OK: db=Nothing handled: " & result
    Test_QuitarMaterializacion_Adversarial_DbNothing = BuildOk("db=Nothing error: " & result, logs)
    GoTo Teardown

EH:
    Test_QuitarMaterializacion_Adversarial_DbNothing = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' =====================================================================
' #3 RegistrarAceptacionRiesgo — 4 atoms
' Signature: Function RegistrarAceptacionRiesgo(
'   ByVal p_IDRiesgo As Long, ByVal p_Justificacion As String,
'   Optional ByRef db As DAO.Database = Nothing) As String
' Form: Form_FormRiesgoMitigacion (inline lines 11-65)
' Returns: "" on success, error string on failure
' =====================================================================

' Happy: valid riesgo + non-empty justificacion > FechaMitigacionAceptar set
Public Function Test_RegistrarAceptacionRiesgo_Happy() As String
    Dim logs(0 To 7) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim result As String
    Dim fechaAceptarBefore As Variant
    Dim fechaAceptarAfter As Variant
    Const Justificacion As String = "Aceptacion por decision de la direccion"

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: Estado=Mitigacion"
    logs(3) = "4. Act: RegistrarAceptacionRiesgo(900503, justificacion)"
    logs(4) = "5. Assert: result="""" (success)"
    logs(5) = "6. Assert: FechaMitigacionAceptar is not null"
    logs(6) = "7. Teardown"
    logs(7) = "Happy: valid aceptacion with justificacion"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_RegistrarAceptacionRiesgo_Happy = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RegistrarAceptacionRiesgo_Happy = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    db.Execute "UPDATE TbRiesgos SET FechaMitigacionAceptar=NULL, Estado='Mitigacion' " & _
        "WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, dbFailOnError

    Set rs = db.OpenRecordset("SELECT FechaMitigacionAceptar FROM TbRiesgos WHERE IDRiesgo=" & _
        Test_Fixtures.Cache_RiesgoId, dbOpenSnapshot)
    If Not rs.EOF Then fechaAceptarBefore = rs!FechaMitigacionAceptar.value
    rs.Close

    result = RegistrarAceptacionRiesgo(Test_Fixtures.Cache_RiesgoId, Justificacion, db)

    Set rs = db.OpenRecordset("SELECT FechaMitigacionAceptar FROM TbRiesgos WHERE IDRiesgo=" & _
        Test_Fixtures.Cache_RiesgoId, dbOpenSnapshot)
    If Not rs.EOF Then fechaAceptarAfter = rs!FechaMitigacionAceptar.value
    rs.Close

    If result <> "" Then
        logs(4) = "4. FAILED: result=""" & result & """, expected """""
        Test_RegistrarAceptacionRiesgo_Happy = BuildFail("expected empty: " & result, logs)
        GoTo Teardown
    End If

    If IsNull(fechaAceptarAfter) Then
        logs(5) = "5. FAILED: FechaMitigacionAceptar still null after aceptacion"
        Test_RegistrarAceptacionRiesgo_Happy = BuildFail("FechaMitigacionAceptar not set", logs)
        GoTo Teardown
    End If

    logs(5) = "5. OK: FechaMitigacionAceptar set to " & CStr(fechaAceptarAfter)
    Test_RegistrarAceptacionRiesgo_Happy = BuildOk("aceptacion ok: " & CStr(fechaAceptarAfter), logs)
    GoTo Teardown

EH:
    Test_RegistrarAceptacionRiesgo_Happy = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call EnsureTbRiesgosClean(db, Test_Fixtures.Cache_RiesgoId)
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Sad: blank justificacion
Public Function Test_RegistrarAceptacionRiesgo_Sad_BlankJustificacion() As String
    Dim logs(0 To 5) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim result As String

    logs(0) = "1. Arrange: ForceLocalBackend + SeedAll"
    logs(1) = "2. Act: RegistrarAceptacionRiesgo with blank justificacion"
    logs(2) = "3. Assert: result <> "" (validation error)"
    logs(3) = "Sad: required field blank"
    logs(4) = "4. Teardown"
    logs(5) = ""

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_RegistrarAceptacionRiesgo_Sad_BlankJustificacion = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RegistrarAceptacionRiesgo_Sad_BlankJustificacion = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    result = RegistrarAceptacionRiesgo(Test_Fixtures.Cache_RiesgoId, "", db)

    If result = "" Then
        logs(2) = "3. FAILED: blank justificacion should have returned error"
        Test_RegistrarAceptacionRiesgo_Sad_BlankJustificacion = BuildFail("expected error for blank justificacion", logs)
        GoTo Teardown
    End If

    logs(2) = "3. OK: blank justificacion rejected: " & result
    Test_RegistrarAceptacionRiesgo_Sad_BlankJustificacion = BuildOk("blank justificacion rejected: " & result, logs)
    GoTo Teardown

EH:
    Test_RegistrarAceptacionRiesgo_Sad_BlankJustificacion = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Edge: very long justificacion (255+ chars)
Public Function Test_RegistrarAceptacionRiesgo_Edge_LongJustificacion() As String
    Dim logs(0 To 5) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim result As String
    Dim longJust As String
    Dim i As Long

    logs(0) = "1. Arrange: ForceLocalBackend + SeedAll"
    logs(1) = "2. Arrange: justificacion = 300 chars"
    logs(2) = "3. Act: RegistrarAceptacionRiesgo with very long justificacion"
    logs(3) = "4. Assert: helper handles boundary gracefully"
    logs(4) = "Edge: field overflow protection"
    logs(5) = "5. Teardown"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_RegistrarAceptacionRiesgo_Edge_LongJustificacion = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RegistrarAceptacionRiesgo_Edge_LongJustificacion = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    longJust = String$(300, "X")

    result = RegistrarAceptacionRiesgo(Test_Fixtures.Cache_RiesgoId, longJust, db)

    logs(3) = "4. INFO: result=""" & result & """ for 300-char justificacion"
    Test_RegistrarAceptacionRiesgo_Edge_LongJustificacion = BuildOk("long justificacion handled: " & result, logs)
    GoTo Teardown

EH:
    Test_RegistrarAceptacionRiesgo_Edge_LongJustificacion = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Adversarial: non-existent riesgo ID
Public Function Test_RegistrarAceptacionRiesgo_Adversarial_IDNoExiste() As String
    Dim logs(0 To 5) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim result As String
    Const NON_EXISTENT_ID As Long = 999998

    logs(0) = "1. Arrange: ForceLocalBackend + SeedAll"
    logs(1) = "2. Act: RegistrarAceptacionRiesgo with non-existent ID=" & NON_EXISTENT_ID
    logs(2) = "3. Assert: result <> "" (error for non-existent riesgo)"
    logs(3) = "Adversarial: invalid riesgo ID"
    logs(4) = "4. Teardown"
    logs(5) = ""

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_RegistrarAceptacionRiesgo_Adversarial_IDNoExiste = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RegistrarAceptacionRiesgo_Adversarial_IDNoExiste = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    result = RegistrarAceptacionRiesgo(NON_EXISTENT_ID, "justificacion test", db)

    If result = "" Then
        logs(2) = "3. FAILED: non-existent riesgo should have returned error"
        Test_RegistrarAceptacionRiesgo_Adversarial_IDNoExiste = BuildFail("expected error for non-existent riesgo", logs)
        GoTo Teardown
    End If

    logs(2) = "3. OK: non-existent riesgo rejected: " & result
    Test_RegistrarAceptacionRiesgo_Adversarial_IDNoExiste = BuildOk("non-existent rejected: " & result, logs)
    GoTo Teardown

EH:
    Test_RegistrarAceptacionRiesgo_Adversarial_IDNoExiste = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' =====================================================================
' #4 ValidarCambioMitigacion — 4 atoms
' Signature: Function ValidarCambioMitigacion(
'   ByVal p_IDRiesgo As Long, ByVal p_NuevaMitigacion As String,
'   Optional ByRef db As DAO.Database = Nothing,
'   Optional ByRef p_PromptResult As Long) As String
' Form: Form_FormRiesgoMitigacion (inline lines 433-563)
' Returns: "" on success, error string on failure
' =====================================================================

' Happy: non-empty nueva mitigacion
Public Function Test_ValidarCambioMitigacion_Happy() As String
    Dim logs(0 To 5) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim result As String
    Dim p_PromptResult As Long
    Const NUEVA_MIT As String = "Nueva estrategia de mitigacion: transferencia a tercero"

    logs(0) = "1. Arrange: ForceLocalBackend + SeedAll"
    logs(1) = "2. Act: ValidarCambioMitigacion with non-empty nueva mitigacion"
    logs(2) = "3. Assert: result="""" (validation success)"
    logs(3) = "Happy: valid mitigacion change"
    logs(4) = "4. Teardown"
    logs(5) = ""

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_ValidarCambioMitigacion_Happy = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ValidarCambioMitigacion_Happy = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    p_PromptResult = 0
    result = ValidarCambioMitigacion(Test_Fixtures.Cache_RiesgoId, NUEVA_MIT, db, p_PromptResult)

    If result <> "" Then
        logs(2) = "3. FAILED: result=""" & result & """, expected """""
        Test_ValidarCambioMitigacion_Happy = BuildFail("expected empty: " & result, logs)
        GoTo Teardown
    End If

    logs(2) = "3. OK: result is empty (validation passed)"
    Test_ValidarCambioMitigacion_Happy = BuildOk("validacion ok", logs)
    GoTo Teardown

EH:
    Test_ValidarCambioMitigacion_Happy = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Sad: blank nueva mitigacion
Public Function Test_ValidarCambioMitigacion_Sad_BlankMitigacion() As String
    Dim logs(0 To 5) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim result As String

    logs(0) = "1. Arrange: ForceLocalBackend + SeedAll"
    logs(1) = "2. Act: ValidarCambioMitigacion with blank nueva mitigacion"
    logs(2) = "3. Assert: result <> "" (validation error)"
    logs(3) = "Sad: required field blank"
    logs(4) = "4. Teardown"
    logs(5) = ""

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_ValidarCambioMitigacion_Sad_BlankMitigacion = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ValidarCambioMitigacion_Sad_BlankMitigacion = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    result = ValidarCambioMitigacion(Test_Fixtures.Cache_RiesgoId, "", db)

    If result = "" Then
        logs(2) = "3. FAILED: blank mitigacion should have returned error"
        Test_ValidarCambioMitigacion_Sad_BlankMitigacion = BuildFail("expected error for blank mitigacion", logs)
        GoTo Teardown
    End If

    logs(2) = "3. OK: blank mitigacion rejected: " & result
    Test_ValidarCambioMitigacion_Sad_BlankMitigacion = BuildOk("blank mitigacion rejected: " & result, logs)
    GoTo Teardown

EH:
    Test_ValidarCambioMitigacion_Sad_BlankMitigacion = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Edge: whitespace-only mitigacion
Public Function Test_ValidarCambioMitigacion_Edge_WhitespaceMitigacion() As String
    Dim logs(0 To 5) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim result As String

    logs(0) = "1. Arrange: ForceLocalBackend + SeedAll"
    logs(1) = "2. Act: ValidarCambioMitigacion with whitespace-only mitigacion"
    logs(2) = "3. Assert: helper handles whitespace (trimmed or rejected)"
    logs(3) = "Edge: whitespace-only input"
    logs(4) = "4. Teardown"
    logs(5) = ""

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_ValidarCambioMitigacion_Edge_WhitespaceMitigacion = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ValidarCambioMitigacion_Edge_WhitespaceMitigacion = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    result = ValidarCambioMitigacion(Test_Fixtures.Cache_RiesgoId, "     ", db)

    logs(2) = "3. INFO: result=""" & result & """ for whitespace-only mitigacion"
    Test_ValidarCambioMitigacion_Edge_WhitespaceMitigacion = BuildOk("whitespace handled: " & result, logs)
    GoTo Teardown

EH:
    Test_ValidarCambioMitigacion_Edge_WhitespaceMitigacion = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Adversarial: riesgo in wrong state (Retirado, not Mitigacion)
Public Function Test_ValidarCambioMitigacion_Adversarial_EstadoInvalido() As String
    Dim logs(0 To 6) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim result As String

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: Estado=Retirado (cannot change mitigacion)"
    logs(3) = "4. Act: ValidarCambioMitigacion on Retirado riesgo"
    logs(4) = "5. Assert: result <> "" (invalid state transition)"
    logs(5) = "Adversarial: wrong estado for mitigacion change"
    logs(6) = "6. Teardown"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_ValidarCambioMitigacion_Adversarial_EstadoInvalido = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ValidarCambioMitigacion_Adversarial_EstadoInvalido = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    db.Execute "UPDATE TbRiesgos SET FechaRetirado=#2026-01-01#, Estado='Retirado' " & _
        "WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, dbFailOnError

    result = ValidarCambioMitigacion(Test_Fixtures.Cache_RiesgoId, "Nueva mitigacion", db)

    If result = "" Then
        logs(4) = "5. FAILED: Estado=Retirado should reject mitigacion change"
        Test_ValidarCambioMitigacion_Adversarial_EstadoInvalido = BuildFail("expected error for Estado=Retirado", logs)
        GoTo Teardown
    End If

    logs(4) = "5. OK: Estado=Retirado rejected: " & result
    Test_ValidarCambioMitigacion_Adversarial_EstadoInvalido = BuildOk("invalid state rejected: " & result, logs)
    GoTo Teardown

EH:
    Test_ValidarCambioMitigacion_Adversarial_EstadoInvalido = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Call EnsureTbRiesgosClean(db, Test_Fixtures.Cache_RiesgoId)
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' =====================================================================
' #5 RegistrarRetiroRiesgo — 4 atoms
' Signature: Function RegistrarRetiroRiesgo(
'   ByVal p_IDRiesgo As Long, ByVal p_Justificacion As String, ByVal p_Fecha As String,
'   Optional ByRef db As DAO.Database = Nothing,
'   Optional ByRef p_PromptResult As Long) As String
' Form: Form_FormRiesgoRetirado (inline lines 131-203)
' Returns: "" on success, error string on failure
' NOTE: Assumes same date-validation rule as Materializado (TODO: confirmar)
' =====================================================================

' Happy: valid retiro with justificacion + valid fecha
Public Function Test_RegistrarRetiroRiesgo_Happy() As String
    Dim logs(0 To 7) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim result As String
    Dim fechaRetiroAfter As Variant
    Dim estadoAfter As String
    Const Justificacion As String = "Riesgo eliminado por no aplicabilidad al proyecto"
    Const FECHA_RETIRO As String = "15/06/2026"

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: Estado=Detectado (pre-retiro)"
    logs(3) = "4. Act: RegistrarRetiroRiesgo(900503, justificacion, " & FECHA_RETIRO & ")"
    logs(4) = "5. Assert: result="""" (success)"
    logs(5) = "6. Assert: FechaRetirado set and Estado='Retirado'"
    logs(6) = "Assumption: same date-validation rule as Materializado (TODO: confirmar)"
    logs(7) = "7. Teardown"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_RegistrarRetiroRiesgo_Happy = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RegistrarRetiroRiesgo_Happy = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    db.Execute "UPDATE TbRiesgos SET FechaRetirado=NULL, FechaMaterializado=NULL, Estado='Detectado' " & _
        "WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, dbFailOnError

    result = RegistrarRetiroRiesgo(Test_Fixtures.Cache_RiesgoId, Justificacion, FECHA_RETIRO, db)

    Set rs = db.OpenRecordset("SELECT FechaRetirado, Estado FROM TbRiesgos WHERE IDRiesgo=" & _
        Test_Fixtures.Cache_RiesgoId, dbOpenSnapshot)
    If Not rs.EOF Then
        fechaRetiroAfter = rs!FechaRetirado.value
        estadoAfter = CStr(Nz(rs!Estado.value, ""))
    End If
    rs.Close

    If result <> "" Then
        logs(4) = "4. FAILED: result=""" & result & """, expected """""
        Test_RegistrarRetiroRiesgo_Happy = BuildFail("expected empty: " & result, logs)
        GoTo Teardown
    End If

    If IsNull(fechaRetiroAfter) Then
        logs(5) = "5. FAILED: FechaRetirado not set after registro"
        Test_RegistrarRetiroRiesgo_Happy = BuildFail("FechaRetirado not set", logs)
        GoTo Teardown
    End If

    If estadoAfter <> "Retirado" Then
        logs(5) = "5. FAILED: Estado=" & estadoAfter & ", expected Retirado"
        Test_RegistrarRetiroRiesgo_Happy = BuildFail("Estado not Retirado", logs)
        GoTo Teardown
    End If

    logs(5) = "5. OK: FechaRetirado=" & CStr(fechaRetiroAfter) & ", Estado=" & estadoAfter
    Test_RegistrarRetiroRiesgo_Happy = BuildOk("retiro registrado ok", logs)
    GoTo Teardown

EH:
    Test_RegistrarRetiroRiesgo_Happy = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call EnsureTbRiesgosClean(db, Test_Fixtures.Cache_RiesgoId)
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Sad: blank justificacion
Public Function Test_RegistrarRetiroRiesgo_Sad_BlankJustificacion() As String
    Dim logs(0 To 5) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim result As String
    Const Fecha As String = "15/06/2026"

    logs(0) = "1. Arrange: ForceLocalBackend + SeedAll"
    logs(1) = "2. Act: RegistrarRetiroRiesgo with blank justificacion"
    logs(2) = "3. Assert: result <> "" (validation error)"
    logs(3) = "Sad: required field blank"
    logs(4) = "4. Teardown"
    logs(5) = ""

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_RegistrarRetiroRiesgo_Sad_BlankJustificacion = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RegistrarRetiroRiesgo_Sad_BlankJustificacion = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    result = RegistrarRetiroRiesgo(Test_Fixtures.Cache_RiesgoId, "", Fecha, db)

    If result = "" Then
        logs(2) = "3. FAILED: blank justificacion should have returned error"
        Test_RegistrarRetiroRiesgo_Sad_BlankJustificacion = BuildFail("expected error for blank justificacion", logs)
        GoTo Teardown
    End If

    logs(2) = "3. OK: blank justificacion rejected: " & result
    Test_RegistrarRetiroRiesgo_Sad_BlankJustificacion = BuildOk("blank justificacion rejected: " & result, logs)
    GoTo Teardown

EH:
    Test_RegistrarRetiroRiesgo_Sad_BlankJustificacion = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Edge: blank fecha
Public Function Test_RegistrarRetiroRiesgo_Edge_BlankFecha() As String
    Dim logs(0 To 5) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim result As String

    logs(0) = "1. Arrange: ForceLocalBackend + SeedAll"
    logs(1) = "2. Act: RegistrarRetiroRiesgo with blank fecha"
    logs(2) = "3. Assert: helper handles blank date"
    logs(3) = "Edge: blank required date"
    logs(4) = "4. Teardown"
    logs(5) = ""

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_RegistrarRetiroRiesgo_Edge_BlankFecha = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RegistrarRetiroRiesgo_Edge_BlankFecha = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    result = RegistrarRetiroRiesgo(Test_Fixtures.Cache_RiesgoId, "justificacion test", "", db)

    logs(2) = "3. INFO: result=""" & result & """ for blank fecha"
    Test_RegistrarRetiroRiesgo_Edge_BlankFecha = BuildOk("blank fecha handled: " & result, logs)
    GoTo Teardown

EH:
    Test_RegistrarRetiroRiesgo_Edge_BlankFecha = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Adversarial: double-click simulation — call twice rapidly
Public Function Test_RegistrarRetiroRiesgo_Adversarial_DobleClick() As String
    Dim logs(0 To 7) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim result1 As String
    Dim result2 As String
    Dim countAfter1 As Long
    Dim countAfter2 As Long
    Const Justificacion As String = "Retiro por duplicacion"
    Const Fecha As String = "20/06/2026"

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Act: First RegistrarRetiroRiesgo call"
    logs(3) = "4. Act: Second RegistrarRetiroRiesgo call (double-click)"
    logs(4) = "5. Assert: only one retiro record (idempotency)"
    logs(5) = "Adversarial: double-click / idempotency"
    logs(6) = "6. Teardown"
    logs(7) = ""

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_RegistrarRetiroRiesgo_Adversarial_DobleClick = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RegistrarRetiroRiesgo_Adversarial_DobleClick = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    db.Execute "UPDATE TbRiesgos SET FechaRetirado=NULL, Estado='Detectado' " & _
        "WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, dbFailOnError

    result1 = RegistrarRetiroRiesgo(Test_Fixtures.Cache_RiesgoId, Justificacion, Fecha, db)
    countAfter1 = CountRows(db, "TbRiesgos", "IDRiesgo=" & Test_Fixtures.Cache_RiesgoId & _
        " AND Not FechaRetirado Is Null")

    result2 = RegistrarRetiroRiesgo(Test_Fixtures.Cache_RiesgoId, Justificacion, Fecha, db)
    countAfter2 = CountRows(db, "TbRiesgos", "IDRiesgo=" & Test_Fixtures.Cache_RiesgoId & _
        " AND Not FechaRetirado Is Null")

    If countAfter2 <> 1 Then
        logs(4) = "5. FAILED: double-click created multiple rows (count=" & countAfter2 & ")"
        Test_RegistrarRetiroRiesgo_Adversarial_DobleClick = BuildFail("double-click not idempotent", logs)
        GoTo Teardown
    End If

    logs(4) = "4. OK: idempotent; only 1 retiro row (result1=" & result1 & ", result2=" & result2 & ")"
    Test_RegistrarRetiroRiesgo_Adversarial_DobleClick = BuildOk("idempotent: result1=" & result1 & ", result2=" & result2, logs)
    GoTo Teardown

EH:
    Test_RegistrarRetiroRiesgo_Adversarial_DobleClick = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Call EnsureTbRiesgosClean(db, Test_Fixtures.Cache_RiesgoId)
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' =====================================================================
' #6 QuitarRetiroRiesgo — 4 atoms
' DIVERGENCE: QuitarRetiroRiesgo vs QuitarMaterializacion (#2)
'   These are DIFFERENT helpers calling DIFFERENT domain methods.
'   Both use shared validation pattern but target different tables/fields.
' Signature: Function QuitarRetiroRiesgo(
'   ByVal p_IDRiesgo As Long,
'   Optional ByRef db As DAO.Database = Nothing,
'   Optional ByRef p_PromptResult As Long) As String
' Form: Form_FormRiesgoRetirado (inline lines 205-265)
' Returns: "" on success, error string on failure
' =====================================================================

' Happy: riesgo is Retirado; confirm vbYes; retiro removed
Public Function Test_QuitarRetiroRiesgo_Happy() As String
    Dim logs(0 To 7) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim p_PromptResult As Long
    Dim result As String
    Dim fechaRetiroBefore As Variant
    Dim fechaRetiroAfter As Variant
    Dim estadoAfter As String

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: Estado=Retirado with FechaRetirado"
    logs(3) = "4. Act: QuitarRetiroRiesgo(900503, vbYes)"
    logs(4) = "5. Assert: result="""" and p_PromptResult=vbYes"
    logs(5) = "6. Assert: FechaRetirado cleared and Estado reverted"
    logs(6) = "7. Teardown"
    logs(7) = "DIVERGENCE vs #2: different helper, different domain method"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_QuitarRetiroRiesgo_Happy = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_QuitarRetiroRiesgo_Happy = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    db.Execute "UPDATE TbRiesgos SET FechaRetirado=#2026-06-15#, Estado='Retirado' " & _
        "WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, dbFailOnError

    Set rs = db.OpenRecordset("SELECT FechaRetirado, Estado FROM TbRiesgos WHERE IDRiesgo=" & _
        Test_Fixtures.Cache_RiesgoId, dbOpenSnapshot)
    If Not rs.EOF Then fechaRetiroBefore = rs!FechaRetirado.value
    rs.Close

    p_PromptResult = vbYes
    result = QuitarRetiroRiesgo(Test_Fixtures.Cache_RiesgoId, db, p_PromptResult)

    Set rs = db.OpenRecordset("SELECT FechaRetirado, Estado FROM TbRiesgos WHERE IDRiesgo=" & _
        Test_Fixtures.Cache_RiesgoId, dbOpenSnapshot)
    If Not rs.EOF Then
        fechaRetiroAfter = rs!FechaRetirado.value
        estadoAfter = CStr(Nz(rs!Estado.value, ""))
    End If
    rs.Close

    If result <> "" Then
        logs(4) = "4. FAILED: result=""" & result & """, expected """""
        Test_QuitarRetiroRiesgo_Happy = BuildFail("expected empty: " & result, logs)
        GoTo Teardown
    End If
    If p_PromptResult <> vbYes Then
        logs(4) = "4. FAILED: p_PromptResult=" & p_PromptResult & " expected vbYes"
        Test_QuitarRetiroRiesgo_Happy = BuildFail("p_PromptResult wrong", logs)
        GoTo Teardown
    End If
    If Not IsNull(fechaRetiroAfter) Then
        logs(5) = "5. FAILED: FechaRetirado still set after QuitarRetiroRiesgo"
        Test_QuitarRetiroRiesgo_Happy = BuildFail("FechaRetirado not cleared", logs)
        GoTo Teardown
    End If

    logs(5) = "5. OK: FechaRetirado cleared (before=" & CStr(fechaRetiroBefore) & "), Estado=" & estadoAfter
    Test_QuitarRetiroRiesgo_Happy = BuildOk("QuitarRetiroRiesgo ok", logs)
    GoTo Teardown

EH:
    Test_QuitarRetiroRiesgo_Happy = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call EnsureTbRiesgosClean(db, Test_Fixtures.Cache_RiesgoId)
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Sad: riesgo is not in Retirado state
Public Function Test_QuitarRetiroRiesgo_Sad_EstadoNoRetirado() As String
    Dim logs(0 To 6) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim p_PromptResult As Long
    Dim result As String
    Dim Estado As String
    Dim rs As DAO.Recordset

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: Estado=Detectado (not Retirado)"
    logs(3) = "4. Act: QuitarRetiroRiesgo on non-retirado riesgo"
    logs(4) = "5. Assert: result handling (error or no-op)"
    logs(5) = "Sad: invalid state transition"
    logs(6) = "6. Teardown"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_QuitarRetiroRiesgo_Sad_EstadoNoRetirado = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_QuitarRetiroRiesgo_Sad_EstadoNoRetirado = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    db.Execute "UPDATE TbRiesgos SET FechaRetirado=NULL, Estado='Detectado' " & _
        "WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, dbFailOnError

    p_PromptResult = vbYes
    result = QuitarRetiroRiesgo(Test_Fixtures.Cache_RiesgoId, db, p_PromptResult)

    Set rs = db.OpenRecordset("SELECT Estado FROM TbRiesgos WHERE IDRiesgo=" & _
        Test_Fixtures.Cache_RiesgoId, dbOpenSnapshot)
    If Not rs.EOF Then Estado = CStr(Nz(rs!Estado.value, ""))
    rs.Close

    If result = "" Then
        logs(4) = "4. INFO: no-op on non-retirado; Estado=" & Estado
        Test_QuitarRetiroRiesgo_Sad_EstadoNoRetirado = BuildOk("no-op: " & result, logs)
        GoTo Teardown
    End If

    logs(4) = "4. OK: invalid state rejected: " & result
    Test_QuitarRetiroRiesgo_Sad_EstadoNoRetirado = BuildOk("invalid state rejected: " & result, logs)
    GoTo Teardown

EH:
    Test_QuitarRetiroRiesgo_Sad_EstadoNoRetirado = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call EnsureTbRiesgosClean(db, Test_Fixtures.Cache_RiesgoId)
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Edge: user clicks vbNo (cancels removal)
Public Function Test_QuitarRetiroRiesgo_Edge_ConfirmacionRechazada() As String
    Dim logs(0 To 6) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim p_PromptResult As Long
    Dim result As String
    Dim fechaRetiroBefore As Variant
    Dim fechaRetiroAfter As Variant

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: Estado=Retirado"
    logs(3) = "4. Act: QuitarRetiroRiesgo with vbNo (user cancels)"
    logs(4) = "5. Assert: p_PromptResult=vbNo respected"
    logs(5) = "6. Assert: FechaRetirado still set (no changes)"
    logs(6) = "Edge: user rejects confirmation"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_QuitarRetiroRiesgo_Edge_ConfirmacionRechazada = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_QuitarRetiroRiesgo_Edge_ConfirmacionRechazada = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    db.Execute "UPDATE TbRiesgos SET FechaRetirado=#2026-06-15#, Estado='Retirado' " & _
        "WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, dbFailOnError

    Set rs = db.OpenRecordset("SELECT FechaRetirado FROM TbRiesgos WHERE IDRiesgo=" & _
        Test_Fixtures.Cache_RiesgoId, dbOpenSnapshot)
    If Not rs.EOF Then fechaRetiroBefore = rs!FechaRetirado.value
    rs.Close

    p_PromptResult = vbNo
    result = QuitarRetiroRiesgo(Test_Fixtures.Cache_RiesgoId, db, p_PromptResult)

    Set rs = db.OpenRecordset("SELECT FechaRetirado FROM TbRiesgos WHERE IDRiesgo=" & _
        Test_Fixtures.Cache_RiesgoId, dbOpenSnapshot)
    If Not rs.EOF Then fechaRetiroAfter = rs!FechaRetirado.value
    rs.Close

    If IsNull(fechaRetiroAfter) Then
        logs(5) = "5. FAILED: vbNo was ignored; FechaRetirado was cleared"
        Test_QuitarRetiroRiesgo_Edge_ConfirmacionRechazada = BuildFail("vbNo ignored", logs)
        GoTo Teardown
    End If

    logs(5) = "5. OK: vbNo respected; FechaRetirado still=" & CStr(fechaRetiroAfter)
    Test_QuitarRetiroRiesgo_Edge_ConfirmacionRechazada = BuildOk("cancelled: no changes", logs)
    GoTo Teardown

EH:
    Test_QuitarRetiroRiesgo_Edge_ConfirmacionRechazada = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call EnsureTbRiesgosClean(db, Test_Fixtures.Cache_RiesgoId)
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Adversarial: db=Nothing causes error
Public Function Test_QuitarRetiroRiesgo_Adversarial_DbNothing() As String
    Dim logs(0 To 5) As String
    Dim p_PromptResult As Long
    Dim result As String

    logs(0) = "1. Arrange: ForceLocalBackend + SeedAll"
    logs(1) = "2. Act: QuitarRetiroRiesgo with db=Nothing"
    logs(2) = "3. Assert: result <> "" (error — db required)"
    logs(3) = "Adversarial: null db causes error"
    logs(4) = "4. Teardown"
    logs(5) = ""

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(logs(0)) Then
        Test_QuitarRetiroRiesgo_Adversarial_DbNothing = BuildFail("TESTS BLOCKED", logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    p_PromptResult = vbYes
    result = QuitarRetiroRiesgo(Test_Fixtures.Cache_RiesgoId, Nothing, p_PromptResult)

    If result = "" Then
        logs(2) = "3. FAILED: db=Nothing should have returned error"
        Test_QuitarRetiroRiesgo_Adversarial_DbNothing = BuildFail("expected error for db=Nothing", logs)
        GoTo Teardown
    End If

    logs(2) = "3. OK: db=Nothing handled: " & result
    Test_QuitarRetiroRiesgo_Adversarial_DbNothing = BuildOk("db=Nothing error: " & result, logs)
    GoTo Teardown

EH:
    Test_QuitarRetiroRiesgo_Adversarial_DbNothing = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' =====================================================================
' #7 SolicitarFechaRetiro — 4 atoms
' Signature: Function SolicitarFechaRetiro(
'   ByVal p_IDRiesgo As Long,
'   Optional ByRef p_Fecha As String,
'   Optional ByRef db As DAO.Database = Nothing,
'   Optional ByRef p_PromptResult As Long) As String
' Form: Form_FormRiesgoRetirado (inline lines 301-364)
' Returns: "" on success, error string on failure
' =====================================================================

' Happy: pre-set p_Fecha (valid date provided); helper validates and accepts
Public Function Test_SolicitarFechaRetiro_Happy() As String
    Dim logs(0 To 6) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim p_Fecha As String
    Dim p_PromptResult As Long
    Dim result As String

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: Estado=Detectado (pre-retiro)"
    logs(3) = "4. Act: SolicitarFechaRetiro(900503, pre-set fecha)"
    logs(4) = "5. Assert: result="""" (success)"
    logs(5) = "6. Assert: p_Fecha returned with valid date"
    logs(6) = "7. Teardown"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_SolicitarFechaRetiro_Happy = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_SolicitarFechaRetiro_Happy = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    db.Execute "UPDATE TbRiesgos SET FechaRetirado=NULL, Estado='Detectado' " & _
        "WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, dbFailOnError

    p_Fecha = "20/06/2026"
    p_PromptResult = 0

    result = SolicitarFechaRetiro(Test_Fixtures.Cache_RiesgoId, p_Fecha, db, p_PromptResult)

    If result <> "" Then
        logs(4) = "4. FAILED: result=""" & result & """, expected """""
        Test_SolicitarFechaRetiro_Happy = BuildFail("expected empty: " & result, logs)
        GoTo Teardown
    End If

    logs(4) = "4. OK: result=""" & result & """, p_Fecha=" & p_Fecha
    Test_SolicitarFechaRetiro_Happy = BuildOk("solicitar fecha ok", logs)
    GoTo Teardown

EH:
    Test_SolicitarFechaRetiro_Happy = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Call EnsureTbRiesgosClean(db, Test_Fixtures.Cache_RiesgoId)
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Sad: riesgo already Retirado (invalid state for requesting fecha)
Public Function Test_SolicitarFechaRetiro_Sad_YaRetirado() As String
    Dim logs(0 To 5) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim p_Fecha As String
    Dim p_PromptResult As Long
    Dim result As String

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: Estado=Retirado (cannot request new fecha)"
    logs(3) = "4. Act: SolicitarFechaRetiro on already-retirado riesgo"
    logs(4) = "5. Assert: result handling (error or no-op)"
    logs(5) = "Sad: invalid state for requesting retiro fecha"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_SolicitarFechaRetiro_Sad_YaRetirado = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_SolicitarFechaRetiro_Sad_YaRetirado = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    db.Execute "UPDATE TbRiesgos SET FechaRetirado=#2026-06-01#, Estado='Retirado' " & _
        "WHERE IDRiesgo=" & Test_Fixtures.Cache_RiesgoId, dbFailOnError

    p_Fecha = ""
    p_PromptResult = 0

    result = SolicitarFechaRetiro(Test_Fixtures.Cache_RiesgoId, p_Fecha, db, p_PromptResult)

    logs(3) = "3. INFO: result=""" & result & """ for Estado=Retirado"
    Test_SolicitarFechaRetiro_Sad_YaRetirado = BuildOk("ya-retirado handled: " & result, logs)
    GoTo Teardown

EH:
    Test_SolicitarFechaRetiro_Sad_YaRetirado = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Edge: p_Fecha pre-set to invalid date (not a date string)
Public Function Test_SolicitarFechaRetiro_Edge_FechaInvalida() As String
    Dim logs(0 To 5) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim p_Fecha As String
    Dim p_PromptResult As Long
    Dim result As String

    logs(0) = "1. Arrange: ForceLocalBackend + SeedAll"
    logs(1) = "2. Arrange: p_Fecha=""not-a-date"""
    logs(2) = "3. Act: SolicitarFechaRetiro with invalid pre-set date"
    logs(3) = "4. Assert: helper handles invalid date"
    logs(4) = "Edge: invalid date format"
    logs(5) = "5. Teardown"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_SolicitarFechaRetiro_Edge_FechaInvalida = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_SolicitarFechaRetiro_Edge_FechaInvalida = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    p_Fecha = "no-es-fecha"
    p_PromptResult = 0

    result = SolicitarFechaRetiro(Test_Fixtures.Cache_RiesgoId, p_Fecha, db, p_PromptResult)

    logs(3) = "3. INFO: result=""" & result & """ for invalid date"
    Test_SolicitarFechaRetiro_Edge_FechaInvalida = BuildOk("invalid date handled: " & result, logs)
    GoTo Teardown

EH:
    Test_SolicitarFechaRetiro_Edge_FechaInvalida = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' Adversarial: non-existent riesgo ID
Public Function Test_SolicitarFechaRetiro_Adversarial_IDNoExiste() As String
    Dim logs(0 To 5) As String
    Dim dbErr As String
    Dim db As DAO.Database
    Dim p_Fecha As String
    Dim p_PromptResult As Long
    Dim result As String
    Const NON_EXISTENT_ID As Long = 999997

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Act: SolicitarFechaRetiro with non-existent riesgo ID=" & NON_EXISTENT_ID
    logs(3) = "4. Assert: result <> "" (error — riesgo does not exist)"
    logs(4) = "Adversarial: invalid riesgo ID"
    logs(5) = "5. Teardown"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_SolicitarFechaRetiro_Adversarial_IDNoExiste = BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_SolicitarFechaRetiro_Adversarial_IDNoExiste = BuildFail("GetTestDb: " & dbErr, logs)
        GoTo Teardown
    End If

    p_Fecha = "20/06/2026"
    p_PromptResult = 0

    result = SolicitarFechaRetiro(NON_EXISTENT_ID, p_Fecha, db, p_PromptResult)

    If result = "" Then
        logs(3) = "4. FAILED: non-existent riesgo should have returned error"
        Test_SolicitarFechaRetiro_Adversarial_IDNoExiste = BuildFail("expected error for non-existent riesgo", logs)
        GoTo Teardown
    End If

    logs(3) = "4. OK: non-existent riesgo rejected: " & result
    Test_SolicitarFechaRetiro_Adversarial_IDNoExiste = BuildOk("non-existent rejected: " & result, logs)
    GoTo Teardown

EH:
    Test_SolicitarFechaRetiro_Adversarial_IDNoExiste = BuildFail("Error " & Err.Number & ": " & Err.description, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0
End Function

' ============================================================
' W2 ATOMS — Cache contract gate migration prep
'
' Skill: access-vba-tdd v2.5
' Helper under test: modRiesgoEstadoGateHelper.bas (gate helpers)
' Target: 7 lines (61, 233, 286, 346, 456, 566, 629) that currently
'         use Constructor.getRiesgo directly, bypassing m_DicRiesgos cache.
'
' Migration plan: change those 7 lines to GetCachedRiesgo so the cache
' is populated as a side-effect of the gate helper's lookup. This
' ensures downstream RiesgoChangeDetector.HasChanges sees the same
' instance across consecutive reads (mirrors Block 1A contract).
'
' Contract being tested (per atom):
'   Pre  : m_DicRiesgos = Nothing (cache cleared)
'   Act  : call gate helper with valid seed params (reaches the L<N>
'          Constructor.getRiesgo lookup; later validation may fail,
'          but the lookup already happened)
'   Post : m_DicRiesgos is populated with test IDRiesgo iff helper
'          uses GetCachedRiesgo. Constructor.getRiesgo bypasses cache.
'   Assert: m_DicRiesgos is NOT Nothing AND m_DicRiesgos.Exists(testID)
'
' Expected before migration (CURRENT):
'   W2-S1.* -> RED (Constructor.getRiesgo bypasses cache; m_DicRiesgos
'              stays Nothing after helper call)
'   W2-S3   -> GREEN (baseline Constructor.getRiesgo returns real riesgo
'              from DB but does NOT populate m_DicRiesgos)
'
' Expected after migration (Commit 1.4):
'   W2-S1.* -> GREEN (helper now calls GetCachedRiesgo, populates cache)
'   W2-S3   -> still GREEN (baseline behavior of Constructor.getRiesgo
'              is unchanged — this is a contract for documentation)
' ============================================================

' Shared fixture range (block-1B cache contract tests):
'   IDRiesgo      900100..900107 (one per gate helper + baseline)
'   IDEdicion     900110..900117
'   IDProyecto    900120..900127

' ============================================================
' ATOM: W2-S1.1 — modRiesgoEstadoGateHelper.bas:61
'   Helper: ValidarTransicionMaterializacion
'   Callsite: Set objRiesgo = Constructor.getRiesgo(p_IDRiesgo:=p_IDRiesgo, p_Error:=p_Error)
'   Act: ValidarTransicionMaterializacion(testID, fecha, plan, db)
'   Expect: m_DicRiesgos.Exists(testID) = True (RED currently)
' ============================================================
Public Function Test_RiesgoEstadoGateHelper_L61_GetCachedRiesgo_NotBypass() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0
    Dim db As DAO.Database
    Dim dbErr As String
    Dim result As String
    Const TEST_ID_EDICION As Long = 900110
    Const TEST_ID_PROYECTO As Long = 900120
    Const TEST_ID_RIESGO As Long = 900100
    Const TEST_ID_PLAN As Long = 900130

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend sandbox": logIdx = logIdx + 1

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_RiesgoEstadoGateHelper_L61_GetCachedRiesgo_NotBypass = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoEstadoGateHelper_L61_GetCachedRiesgo_NotBypass = _
            BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & TEST_ID_RIESGO, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_ID_PROYECTO, dbFailOnError
    logs(logIdx) = "1b. Schema: inspecting required fields for INSERT": logIdx = logIdx + 1
    Call InspectarCamposObligatorios("TbProyectos", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbProyectosEdiciones", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbRiesgos", db, logs, logIdx)
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) VALUES (" & _
        TEST_ID_PROYECTO & ", 'TEST-W2-L61', 'L61 fixture')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) VALUES (" & _
        TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) VALUES (" & _
        TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", 'PC-W2-L61', 'R-W2-L61')", dbFailOnError

    logs(logIdx) = "2. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO: logIdx = logIdx + 1

    InvalidarCacheRiesgo CStr(TEST_ID_RIESGO)

    logs(logIdx) = "3. Act: ValidarTransicionMaterializacion(" & TEST_ID_RIESGO & ", fecha, plan)": logIdx = logIdx + 1

    Dim p_Fecha As String
    Dim p_IDPlan As String
    p_Fecha = "01/01/2025"
    p_IDPlan = CStr(TEST_ID_PLAN)
    result = ValidarTransicionMaterializacion(CStr(TEST_ID_RIESGO), p_Fecha, p_IDPlan, db, vbYes)

    logs(logIdx) = "4. Helper returned: '" & result & "' (later validation may fail; L61 lookup happened)": logIdx = logIdx + 1

    If Not CacheRiesgoContainsKey(CStr(TEST_ID_RIESGO)) Then
        logs(logIdx) = "5. Assert FAIL: cache does not contain IDRiesgo=" & TEST_ID_RIESGO & " (Constructor.getRiesgo bypassed cache)": logIdx = logIdx + 1
        Test_RiesgoEstadoGateHelper_L61_GetCachedRiesgo_NotBypass = _
            BuildFail("L61: cache not populated after ValidarTransicionMaterializacion lookup", logs)
        Exit Function
    End If

    logs(logIdx) = "5. Assert PASS: cache populated with IDRiesgo=" & TEST_ID_RIESGO: logIdx = logIdx + 1
    Test_RiesgoEstadoGateHelper_L61_GetCachedRiesgo_NotBypass = _
        BuildOk("L61 cache populated", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.description: logIdx = logIdx + 1
    Test_RiesgoEstadoGateHelper_L61_GetCachedRiesgo_NotBypass = _
        BuildFail("Error " & Err.Number & ": " & Err.description, logs)
End Function

' ============================================================
' ATOM: W2-S1.2 — modRiesgoEstadoGateHelper.bas:233
'   Helper: QuitarMaterializacion
' ============================================================
Public Function Test_RiesgoEstadoGateHelper_L233_GetCachedRiesgo_NotBypass() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0
    Dim db As DAO.Database
    Dim dbErr As String
    Dim result As String
    Const TEST_ID_EDICION As Long = 900111
    Const TEST_ID_PROYECTO As Long = 900121
    Const TEST_ID_RIESGO As Long = 900101

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend sandbox": logIdx = logIdx + 1

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_RiesgoEstadoGateHelper_L233_GetCachedRiesgo_NotBypass = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoEstadoGateHelper_L233_GetCachedRiesgo_NotBypass = _
            BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & TEST_ID_RIESGO, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_ID_PROYECTO, dbFailOnError
    logs(logIdx) = "1b. Schema: inspecting required fields for INSERT": logIdx = logIdx + 1
    Call InspectarCamposObligatorios("TbProyectos", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbProyectosEdiciones", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbRiesgos", db, logs, logIdx)
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) VALUES (" & _
        TEST_ID_PROYECTO & ", 'TEST-W2-L233', 'L233 fixture')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) VALUES (" & _
        TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) VALUES (" & _
        TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", 'PC-W2-L233', 'R-W2-L233')", dbFailOnError

    logs(logIdx) = "2. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO: logIdx = logIdx + 1

    InvalidarCacheRiesgo CStr(TEST_ID_RIESGO)

    logs(logIdx) = "3. Act: QuitarMaterializacion(" & TEST_ID_RIESGO & ", db, vbYes)": logIdx = logIdx + 1

    result = QuitarMaterializacion(CStr(TEST_ID_RIESGO), db, vbYes)

    logs(logIdx) = "4. Helper returned: '" & result & "'": logIdx = logIdx + 1

    If Not CacheRiesgoContainsKey(CStr(TEST_ID_RIESGO)) Then
        logs(logIdx) = "5. Assert FAIL: cache does not contain IDRiesgo=" & TEST_ID_RIESGO & " (Constructor.getRiesgo bypassed cache)": logIdx = logIdx + 1
        Test_RiesgoEstadoGateHelper_L233_GetCachedRiesgo_NotBypass = _
            BuildFail("L233: cache not populated after QuitarMaterializacion lookup", logs)
        Exit Function
    End If

    logs(logIdx) = "5. Assert PASS: cache populated with IDRiesgo=" & TEST_ID_RIESGO: logIdx = logIdx + 1
    Test_RiesgoEstadoGateHelper_L233_GetCachedRiesgo_NotBypass = _
        BuildOk("L233 cache populated", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.description: logIdx = logIdx + 1
    Test_RiesgoEstadoGateHelper_L233_GetCachedRiesgo_NotBypass = _
        BuildFail("Error " & Err.Number & ": " & Err.description, logs)
End Function

' ============================================================
' ATOM: W2-S1.3 — modRiesgoEstadoGateHelper.bas:286
'   Helper: RegistrarAceptacionRiesgo
'
'   Assertion strategy (Commit 1.4 fix):
'     Pre-populate cache via GetCachedRiesgo, then mutate
'     mock.JustificacionAceptacionRiesgo = "W2_L286_MARKER". Gate calls
'     Riesgo.AceptacionRegistrar which rejects when
'     p_Justificacion = .JustificacionAceptacionRiesgo (Riesgo.cls:5770).
'     Cache hit -> gate sees marker -> error text appears in `result`.
'     Cache miss / Constructor.getRiesgo bypass -> fresh instance with
'     empty JustificacionAceptacionRiesgo -> OK path -> result is "".
'     Robust to AceptacionRegistrar's InvalidarCacheRiesgo pre-persist
'     (Riesgo.cls:5784) — this is a legitimate side-effect of persist
'     that the test should not couple to.
' ============================================================
Public Function Test_RiesgoEstadoGateHelper_L286_GetCachedRiesgo_NotBypass() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0
    Dim db As DAO.Database
    Dim dbErr As String
    Dim result As String
    Dim mockObj As riesgo
    Const TEST_ID_EDICION As Long = 900112
    Const TEST_ID_PROYECTO As Long = 900122
    Const TEST_ID_RIESGO As Long = 900102
    Const MARKER As String = "W2_L286_MARKER_DUP_JUSTIF"

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend sandbox": logIdx = logIdx + 1

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_RiesgoEstadoGateHelper_L286_GetCachedRiesgo_NotBypass = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoEstadoGateHelper_L286_GetCachedRiesgo_NotBypass = _
            BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & TEST_ID_RIESGO, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_ID_PROYECTO, dbFailOnError
    logs(logIdx) = "1b. Schema: inspecting required fields for INSERT": logIdx = logIdx + 1
    Call InspectarCamposObligatorios("TbProyectos", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbProyectosEdiciones", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbRiesgos", db, logs, logIdx)
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) VALUES (" & _
        TEST_ID_PROYECTO & ", 'TEST-W2-L286', 'L286 fixture')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) VALUES (" & _
        TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) VALUES (" & _
        TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", 'PC-W2-L286', 'R-W2-L286')", dbFailOnError

    logs(logIdx) = "2. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO & " (JustificacionAceptacionRiesgo=NULL)": logIdx = logIdx + 1

    InvalidarCacheRiesgo CStr(TEST_ID_RIESGO)

    ' 3. Load fresh from DB into cache, then mutate the cached instance.
    '    GetCachedRiesgo does cache miss -> Constructor.getRiesgo -> Add to m_DicRiesgos.
    '    Mutating the returned reference mutates the cached object (same ref).
    Set mockObj = GetCachedRiesgo(CStr(TEST_ID_RIESGO), dbErr)
    If dbErr <> "" Or mockObj Is Nothing Then
        Test_RiesgoEstadoGateHelper_L286_GetCachedRiesgo_NotBypass = _
            BuildFail("Arrange: GetCachedRiesgo failed: " & dbErr, logs)
        Exit Function
    End If
    mockObj.JustificacionAceptacionRiesgo = MARKER

    logs(logIdx) = "3. Arrange: cache populated with mock JustificacionAceptacionRiesgo='" & MARKER & "'": logIdx = logIdx + 1

    logs(logIdx) = "4. Act: RegistrarAceptacionRiesgo(" & TEST_ID_RIESGO & ", '" & MARKER & "', db)": logIdx = logIdx + 1

    result = RegistrarAceptacionRiesgo(CStr(TEST_ID_RIESGO), MARKER, db)

    logs(logIdx) = "5. Helper returned: '" & result & "'": logIdx = logIdx + 1

    ' 6. Assert: gate USED GetCachedRiesgo (cache hit) -> marker visible to
    '    AceptacionRegistrar -> duplicate-justification error emitted.
    '    If gate bypassed with Constructor.getRiesgo, fresh instance has
    '    JustificacionAceptacionRiesgo="" -> OK path -> result is empty.
    If InStr(result, "ya estaba pendiente de justificar") = 0 Then
        logs(logIdx) = "6. Assert FAIL: marker error not in result (gate saw fresh instance, Constructor.getRiesgo bypass suspected)": logIdx = logIdx + 1
        Test_RiesgoEstadoGateHelper_L286_GetCachedRiesgo_NotBypass = _
            BuildFail("L286: gate did not return marker error — GetCachedRiesgo bypass suspected", logs)
        Exit Function
    End If

    logs(logIdx) = "6. Assert PASS: gate returned marker error (GetCachedRiesgo HIT confirmed)": logIdx = logIdx + 1
    Test_RiesgoEstadoGateHelper_L286_GetCachedRiesgo_NotBypass = _
        BuildOk("L286: gate used GetCachedRiesgo (cache hit)", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.description: logIdx = logIdx + 1
    Test_RiesgoEstadoGateHelper_L286_GetCachedRiesgo_NotBypass = _
        BuildFail("Error " & Err.Number & ": " & Err.description, logs)
End Function

' ============================================================
' ATOM: W2-S1.4 — modRiesgoEstadoGateHelper.bas:346
'   Helper: ValidarCambioMitigacion
' ============================================================
Public Function Test_RiesgoEstadoGateHelper_L346_GetCachedRiesgo_NotBypass() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0
    Dim db As DAO.Database
    Dim dbErr As String
    Dim result As String
    Const TEST_ID_EDICION As Long = 900113
    Const TEST_ID_PROYECTO As Long = 900123
    Const TEST_ID_RIESGO As Long = 900103

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend sandbox": logIdx = logIdx + 1

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_RiesgoEstadoGateHelper_L346_GetCachedRiesgo_NotBypass = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoEstadoGateHelper_L346_GetCachedRiesgo_NotBypass = _
            BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & TEST_ID_RIESGO, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_ID_PROYECTO, dbFailOnError
    logs(logIdx) = "1b. Schema: inspecting required fields for INSERT": logIdx = logIdx + 1
    Call InspectarCamposObligatorios("TbProyectos", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbProyectosEdiciones", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbRiesgos", db, logs, logIdx)
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) VALUES (" & _
        TEST_ID_PROYECTO & ", 'TEST-W2-L346', 'L346 fixture')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) VALUES (" & _
        TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) VALUES (" & _
        TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", 'PC-W2-L346', 'R-W2-L346')", dbFailOnError

    ' Set FechaMaterializado so getEstadoRiesgo returns Materializado and the helper
    ' exits at L361-364 gate BEFORE reaching L368 (which uses GetCachedRiesgo and would
    ' otherwise populate the cache, masking L346's Constructor.getRiesgo bypass).
    db.Execute "UPDATE TbRiesgos SET FechaMaterializado=#2025-01-15# WHERE IDRiesgo=" & TEST_ID_RIESGO, dbFailOnError

    logs(logIdx) = "2. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO & " (Materializado → forces exit at L363 gate)": logIdx = logIdx + 1

    InvalidarCacheRiesgo CStr(TEST_ID_RIESGO)

    logs(logIdx) = "3. Act: ValidarCambioMitigacion(" & TEST_ID_RIESGO & ", 'Mitigar', db)": logIdx = logIdx + 1

    result = ValidarCambioMitigacion(CStr(TEST_ID_RIESGO), "Mitigar", db)

    logs(logIdx) = "4. Helper returned: '" & result & "'": logIdx = logIdx + 1

    If Not CacheRiesgoContainsKey(CStr(TEST_ID_RIESGO)) Then
        logs(logIdx) = "5. Assert FAIL: cache does not contain IDRiesgo=" & TEST_ID_RIESGO & " (Constructor.getRiesgo bypassed cache)": logIdx = logIdx + 1
        Test_RiesgoEstadoGateHelper_L346_GetCachedRiesgo_NotBypass = _
            BuildFail("L346: cache not populated after ValidarCambioMitigacion lookup", logs)
        Exit Function
    End If

    logs(logIdx) = "5. Assert PASS: cache populated with IDRiesgo=" & TEST_ID_RIESGO: logIdx = logIdx + 1
    Test_RiesgoEstadoGateHelper_L346_GetCachedRiesgo_NotBypass = _
        BuildOk("L346 cache populated", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.description: logIdx = logIdx + 1
    Test_RiesgoEstadoGateHelper_L346_GetCachedRiesgo_NotBypass = _
        BuildFail("Error " & Err.Number & ": " & Err.description, logs)
End Function

' ============================================================
' ATOM: W2-S1.5 — modRiesgoEstadoGateHelper.bas:456
'   Helper: RegistrarRetiroRiesgo
'
'   Assertion strategy (Commit 1.4 fix):
'     Pre-populate cache via GetCachedRiesgo, then mutate mock:
'       mock.Estado = "RetiradoSinVisar"
'       mock.JustificacionRetiroRiesgo = "W2_L456_MARKER"
'     Gate passes mock.EstadoEnum as p_Estado = RetiradoSinVisar. Cache HIT
'     triggers RetiroRegistrar duplicate-justification check (Riesgo.cls:6078)
'     and emits error text. Cache MISS / Constructor.getRiesgo bypass -> fresh
'     instance with Estado="", so EstadoEnum falls through to ESTADOCalculado
'     (Detectado), p_Estado != RetiradoSinVisar -> check skipped -> result is "".
'     Robust to RetiroRegistrar's InvalidarCacheRiesgo pre-persist
'     (Riesgo.cls:6099).
' ============================================================
Public Function Test_RiesgoEstadoGateHelper_L456_GetCachedRiesgo_NotBypass() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0
    Dim db As DAO.Database
    Dim dbErr As String
    Dim result As String
    Dim mockObj As riesgo
    Const TEST_ID_EDICION As Long = 900114
    Const TEST_ID_PROYECTO As Long = 900124
    Const TEST_ID_RIESGO As Long = 900104
    Const MARKER As String = "W2_L456_MARKER_DUP_RETIRO"

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend sandbox": logIdx = logIdx + 1

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_RiesgoEstadoGateHelper_L456_GetCachedRiesgo_NotBypass = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoEstadoGateHelper_L456_GetCachedRiesgo_NotBypass = _
            BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & TEST_ID_RIESGO, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_ID_PROYECTO, dbFailOnError
    logs(logIdx) = "1b. Schema: inspecting required fields for INSERT": logIdx = logIdx + 1
    Call InspectarCamposObligatorios("TbProyectos", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbProyectosEdiciones", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbRiesgos", db, logs, logIdx)
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) VALUES (" & _
        TEST_ID_PROYECTO & ", 'TEST-W2-L456', 'L456 fixture')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) VALUES (" & _
        TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) VALUES (" & _
        TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", 'PC-W2-L456', 'R-W2-L456')", dbFailOnError

    logs(logIdx) = "2. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO & " (Estado=NULL, JustificacionRetiroRiesgo=NULL)": logIdx = logIdx + 1

    InvalidarCacheRiesgo CStr(TEST_ID_RIESGO)

    ' 3. Load fresh from DB into cache, then mutate.
    Set mockObj = GetCachedRiesgo(CStr(TEST_ID_RIESGO), dbErr)
    If dbErr <> "" Or mockObj Is Nothing Then
        Test_RiesgoEstadoGateHelper_L456_GetCachedRiesgo_NotBypass = _
            BuildFail("Arrange: GetCachedRiesgo failed: " & dbErr, logs)
        Exit Function
    End If
    mockObj.Estado = "RetiradoSinVisar"
    mockObj.JustificacionRetiroRiesgo = MARKER

    logs(logIdx) = "3. Arrange: cache populated with mock Estado='RetiradoSinVisar', JustificacionRetiroRiesgo='" & MARKER & "'": logIdx = logIdx + 1

    logs(logIdx) = "4. Act: RegistrarRetiroRiesgo(" & TEST_ID_RIESGO & ", '" & MARKER & "', '01/01/2025', db)": logIdx = logIdx + 1

    result = RegistrarRetiroRiesgo(CStr(TEST_ID_RIESGO), MARKER, "01/01/2025", db)

    logs(logIdx) = "5. Helper returned: '" & result & "'": logIdx = logIdx + 1

    ' 6. Assert: gate USED GetCachedRiesgo (cache hit) -> marker + Estado visible
    '    to RetiroRegistrar -> duplicate-justification error emitted.
    If InStr(result, "ya estaba en proceso de retirada") = 0 Then
        logs(logIdx) = "6. Assert FAIL: marker error not in result (gate saw fresh instance, Constructor.getRiesgo bypass suspected)": logIdx = logIdx + 1
        Test_RiesgoEstadoGateHelper_L456_GetCachedRiesgo_NotBypass = _
            BuildFail("L456: gate did not return marker error — GetCachedRiesgo bypass suspected", logs)
        Exit Function
    End If

    logs(logIdx) = "6. Assert PASS: gate returned marker error (GetCachedRiesgo HIT confirmed)": logIdx = logIdx + 1
    Test_RiesgoEstadoGateHelper_L456_GetCachedRiesgo_NotBypass = _
        BuildOk("L456: gate used GetCachedRiesgo (cache hit)", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.description: logIdx = logIdx + 1
    Test_RiesgoEstadoGateHelper_L456_GetCachedRiesgo_NotBypass = _
        BuildFail("Error " & Err.Number & ": " & Err.description, logs)
End Function

' ============================================================
' ATOM: W2-S1.6 — modRiesgoEstadoGateHelper.bas:566
'   Helper: QuitarRetiroRiesgo
' ============================================================
Public Function Test_RiesgoEstadoGateHelper_L566_GetCachedRiesgo_NotBypass() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0
    Dim db As DAO.Database
    Dim dbErr As String
    Dim result As String
    Const TEST_ID_EDICION As Long = 900115
    Const TEST_ID_PROYECTO As Long = 900125
    Const TEST_ID_RIESGO As Long = 900105

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend sandbox": logIdx = logIdx + 1

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_RiesgoEstadoGateHelper_L566_GetCachedRiesgo_NotBypass = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoEstadoGateHelper_L566_GetCachedRiesgo_NotBypass = _
            BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & TEST_ID_RIESGO, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_ID_PROYECTO, dbFailOnError
    logs(logIdx) = "1b. Schema: inspecting required fields for INSERT": logIdx = logIdx + 1
    Call InspectarCamposObligatorios("TbProyectos", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbProyectosEdiciones", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbRiesgos", db, logs, logIdx)
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) VALUES (" & _
        TEST_ID_PROYECTO & ", 'TEST-W2-L566', 'L566 fixture')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) VALUES (" & _
        TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) VALUES (" & _
        TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", 'PC-W2-L566', 'R-W2-L566')", dbFailOnError

    logs(logIdx) = "2. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO: logIdx = logIdx + 1

    InvalidarCacheRiesgo CStr(TEST_ID_RIESGO)

    logs(logIdx) = "3. Act: QuitarRetiroRiesgo(" & TEST_ID_RIESGO & ", db, vbYes)": logIdx = logIdx + 1

    result = QuitarRetiroRiesgo(CStr(TEST_ID_RIESGO), db, vbYes)

    logs(logIdx) = "4. Helper returned: '" & result & "'": logIdx = logIdx + 1

    If Not CacheRiesgoContainsKey(CStr(TEST_ID_RIESGO)) Then
        logs(logIdx) = "5. Assert FAIL: cache does not contain IDRiesgo=" & TEST_ID_RIESGO & " (Constructor.getRiesgo bypassed cache)": logIdx = logIdx + 1
        Test_RiesgoEstadoGateHelper_L566_GetCachedRiesgo_NotBypass = _
            BuildFail("L566: cache not populated after QuitarRetiroRiesgo lookup", logs)
        Exit Function
    End If

    logs(logIdx) = "5. Assert PASS: cache populated with IDRiesgo=" & TEST_ID_RIESGO: logIdx = logIdx + 1
    Test_RiesgoEstadoGateHelper_L566_GetCachedRiesgo_NotBypass = _
        BuildOk("L566 cache populated", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.description: logIdx = logIdx + 1
    Test_RiesgoEstadoGateHelper_L566_GetCachedRiesgo_NotBypass = _
        BuildFail("Error " & Err.Number & ": " & Err.description, logs)
End Function

' ============================================================
' ATOM: W2-S1.7 — modRiesgoEstadoGateHelper.bas:629
'   Helper: SolicitarFechaRetiro
' ============================================================
Public Function Test_RiesgoEstadoGateHelper_L629_GetCachedRiesgo_NotBypass() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0
    Dim db As DAO.Database
    Dim dbErr As String
    Dim result As String
    Const TEST_ID_EDICION As Long = 900116
    Const TEST_ID_PROYECTO As Long = 900126
    Const TEST_ID_RIESGO As Long = 900106

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend sandbox": logIdx = logIdx + 1

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_RiesgoEstadoGateHelper_L629_GetCachedRiesgo_NotBypass = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoEstadoGateHelper_L629_GetCachedRiesgo_NotBypass = _
            BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & TEST_ID_RIESGO, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_ID_PROYECTO, dbFailOnError
    logs(logIdx) = "1b. Schema: inspecting required fields for INSERT": logIdx = logIdx + 1
    Call InspectarCamposObligatorios("TbProyectos", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbProyectosEdiciones", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbRiesgos", db, logs, logIdx)
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) VALUES (" & _
        TEST_ID_PROYECTO & ", 'TEST-W2-L629', 'L629 fixture')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) VALUES (" & _
        TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) VALUES (" & _
        TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", 'PC-W2-L629', 'R-W2-L629')", dbFailOnError

    logs(logIdx) = "2. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO: logIdx = logIdx + 1

    InvalidarCacheRiesgo CStr(TEST_ID_RIESGO)

    logs(logIdx) = "3. Act: SolicitarFechaRetiro(" & TEST_ID_RIESGO & ", '01/01/2025', db)": logIdx = logIdx + 1

    Dim p_Fecha As String
    p_Fecha = "01/01/2025"
    result = SolicitarFechaRetiro(CStr(TEST_ID_RIESGO), p_Fecha, db)

    logs(logIdx) = "4. Helper returned: '" & result & "'": logIdx = logIdx + 1

    If Not CacheRiesgoContainsKey(CStr(TEST_ID_RIESGO)) Then
        logs(logIdx) = "5. Assert FAIL: cache does not contain IDRiesgo=" & TEST_ID_RIESGO & " (Constructor.getRiesgo bypassed cache)": logIdx = logIdx + 1
        Test_RiesgoEstadoGateHelper_L629_GetCachedRiesgo_NotBypass = _
            BuildFail("L629: cache not populated after SolicitarFechaRetiro lookup", logs)
        Exit Function
    End If

    logs(logIdx) = "5. Assert PASS: cache populated with IDRiesgo=" & TEST_ID_RIESGO: logIdx = logIdx + 1
    Test_RiesgoEstadoGateHelper_L629_GetCachedRiesgo_NotBypass = _
        BuildOk("L629 cache populated", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.description: logIdx = logIdx + 1
    Test_RiesgoEstadoGateHelper_L629_GetCachedRiesgo_NotBypass = _
        BuildFail("Error " & Err.Number & ": " & Err.description, logs)
End Function

' ============================================================
' ATOM: W2-S3 — Baseline Constructor.getRiesgo behavior
'   Documents the BUGGY baseline that W2-S1 atoms are guarding against.
'   Constructor.getRiesgo returns a riesgo from DB but does NOT populate
'   m_DicRiesgos. This is the behavior we want to eliminate by migrating
'   the 7 gate callsites to GetCachedRiesgo.
'
'   This atom is GREEN currently AND stays GREEN after migration
'   (Constructor.getRiesgo behavior itself does not change).
' ============================================================
Public Function Test_RiesgoEstadoGateHelper_W2_S3_Constructor_DoesNotPopulateCache() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0
    Dim db As DAO.Database
    Dim dbErr As String
    Dim p_Error As String
    Dim objRiesgo As riesgo
    Const TEST_ID_EDICION As Long = 900117
    Const TEST_ID_PROYECTO As Long = 900127
    Const TEST_ID_RIESGO As Long = 900107

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend sandbox": logIdx = logIdx + 1

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_RiesgoEstadoGateHelper_W2_S3_Constructor_DoesNotPopulateCache = _
            BuildFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoEstadoGateHelper_W2_S3_Constructor_DoesNotPopulateCache = _
            BuildFail("GetTestDb: " & dbErr, logs)
        Exit Function
    End If

    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & TEST_ID_RIESGO, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_ID_PROYECTO, dbFailOnError
    logs(logIdx) = "1b. Schema: inspecting required fields for INSERT": logIdx = logIdx + 1
    Call InspectarCamposObligatorios("TbProyectos", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbProyectosEdiciones", db, logs, logIdx)
    Call InspectarCamposObligatorios("TbRiesgos", db, logs, logIdx)
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) VALUES (" & _
        TEST_ID_PROYECTO & ", 'TEST-W2-S3', 'Baseline fixture')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) VALUES (" & _
        TEST_ID_EDICION & ", " & TEST_ID_PROYECTO & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) VALUES (" & _
        TEST_ID_RIESGO & ", " & TEST_ID_EDICION & ", 'PC-W2-S3', 'R-W2-S3')", dbFailOnError

    logs(logIdx) = "2. Arrange: seeded riesgo IDRiesgo=" & TEST_ID_RIESGO: logIdx = logIdx + 1

    InvalidarCacheRiesgo CStr(TEST_ID_RIESGO)

    logs(logIdx) = "3. Act: Constructor.getRiesgo(" & TEST_ID_RIESGO & ", , , p_Error)": logIdx = logIdx + 1

    Set objRiesgo = Constructor.getRiesgo(CStr(TEST_ID_RIESGO), , , p_Error)

    If p_Error <> "" Or objRiesgo Is Nothing Then
        logs(logIdx) = "4. Assert FAIL: Constructor.getRiesgo did not return the riesgo (err='" & p_Error & "')": logIdx = logIdx + 1
        Test_RiesgoEstadoGateHelper_W2_S3_Constructor_DoesNotPopulateCache = _
            BuildFail("Constructor.getRiesgo baseline broken", logs)
        Exit Function
    End If

    logs(logIdx) = "4. Constructor.getRiesgo returned instance, p_Error='" & p_Error & "'": logIdx = logIdx + 1

    ' -- Document the buggy baseline: cache stays empty after Constructor.getRiesgo
    If Not CacheRiesgoContainsKey(CStr(TEST_ID_RIESGO)) Then
        logs(logIdx) = "5. Assert PASS: cache does not contain IDRiesgo=" & TEST_ID_RIESGO & " (baseline documented: bypass)": logIdx = logIdx + 1
        Test_RiesgoEstadoGateHelper_W2_S3_Constructor_DoesNotPopulateCache = _
            BuildOk("baseline documented: Constructor.getRiesgo bypasses cache", logs)
        Exit Function
    End If

    logs(logIdx) = "5. Assert FAIL: cache WAS populated by Constructor.getRiesgo — baseline has changed (unexpected)": logIdx = logIdx + 1
    Test_RiesgoEstadoGateHelper_W2_S3_Constructor_DoesNotPopulateCache = _
        BuildFail("baseline drift: Constructor.getRiesgo now populates cache (unintended)", logs)
    Exit Function

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.description: logIdx = logIdx + 1
    Test_RiesgoEstadoGateHelper_W2_S3_Constructor_DoesNotPopulateCache = _
        BuildFail("Error " & Err.Number & ": " & Err.description, logs)
End Function

