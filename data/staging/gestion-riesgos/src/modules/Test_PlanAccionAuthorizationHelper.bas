Attribute VB_Name = "Test_PlanAccionAuthorizationHelper"
' ============================================================
' Test_PlanAccionAuthorizationHelper — RED atoms for
'   modPlanAccionAuthorizationHelper.ValidarEliminacionAccion
'
' Skill: access-vba-tdd v2.4 + access-vba-e2e-methodology
'
' Scope: 1 helper → 4 scenario classes = 4 atoms
'   Happy      > authorized user + existing action + confirms
'   Sad        > non-existent action; empty ID string
'   Edge       > ID with whitespace / special chars
'   Adversarial> double-click guard (idempotencia); role mismatch
'
' Helper signature (canónica, alineada 2026-06-24):
'   Public Function ValidarEliminacionAccion( _
'       ByVal p_EsMitigacion As Boolean, _
'       ByVal p_IDAccion As String, _
'       Optional ByRef p_PromptResult As Long, _
'       Optional ByRef p_Error As String) As Boolean
'
'   p_PromptResult = vbYes (6) bypasses el MsgBox en Happy/Edge/Adv;
'   p_PromptResult = 0 caería al MsgBox real (prohibido en tests TDD).
'   Sad testea el path "no encontrado", así que el MsgBox nunca llega
'   a evaluarse — igualmente pasamos vbYes por simetría/contrato.
' ============================================================
Option Compare Database
Option Explicit

' ============================================================
' ATOM: Happy — authorized user, existing mitigacion action,
'            user confirms > returns True
'
' Setup:
'   m_ObjUsuarioConectado > authorized user
'   Seed: one valid PM row in TbRiesgosPlanMitigacionPpal
'         with ID >= 900000
' Act: ValidarEliminacionAccion(True, "900001", vbYes)
' Expect: ok=true, value=True
' Teardown: DELETE rows WHERE IDMitigacion BETWEEN 900000 AND 999999
' ============================================================
Public Function Test_ValidarEliminacionAccion_Happy() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    ' -- Arrange --------------------------------------------
    logs(logIdx) = "1. Arrange: ForceLocalBackend"
    logIdx = logIdx + 1

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ValidarEliminacionAccion_Happy = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ValidarEliminacionAccion_Happy = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' -- Seed fixture: PM row ---------------------------------
    ' NOTE: schema-first gate passed via prior inspection of
    ' TbRiesgosPlanMitigacionPpal. PK is IDMitigacion (Long).
    ' Required fields seeded explicitly.
    Const TEST_ID_PM As Long = 900001
    Const TEST_ID_EDICION_PM As Long = 900050
    Const TEST_ID_RIESGO_PM As Long = 900050
    Const TEST_ID_PROYECTO_PM As Long = 900060

    ' Clean any residual from prior runs (FK order: child → parent)
    db.Execute "DELETE FROM TbRiesgosPlanMitigacionPpal WHERE IDMitigacion BETWEEN 900000 AND 999999", _
               dbFailOnError
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900050 AND 900059", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion BETWEEN 900050 AND 900059", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto BETWEEN 900060 AND 900069", dbFailOnError

    ' Seed FK chain: Proyecto → Edicion → Riesgo → PlanMitigacion
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO_PM & ", 'TEST-PM-PROJ', 'Test PM project')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION_PM & ", " & TEST_ID_PROYECTO_PM & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) " & _
               "VALUES (" & TEST_ID_RIESGO_PM & ", " & TEST_ID_EDICION_PM & ", " & _
               "'PC-FIX-900050', 'R-PM-900050')", dbFailOnError

    db.Execute "INSERT INTO TbRiesgosPlanMitigacionPpal " & _
                "(IDMitigacion, DisparadorDelPlan, IDRiesgo) " & _
                "VALUES (" & TEST_ID_PM & ", 'PM test atom', " & TEST_ID_RIESGO_PM & ")", _
               dbFailOnError
    ' Also seed detail row (helper Constructor.getPMAccion queries Detalle, not Ppal)
    db.Execute "INSERT INTO TbRiesgosPlanMitigacionDetalle (IDAccionMitigacion, IDMitigacion, CodAccion, Accion, ResponsableAccion, Estado) " & _
               "VALUES (" & TEST_ID_PM & ", " & TEST_ID_PM & ", 'AC-PM-001', 'Test PM action', 'Test User', 'Pendiente')", _
               dbFailOnError

    logs(logIdx) = "2. Arrange: seeded PM row IDMitigacion=" & TEST_ID_PM
    logIdx = logIdx + 1

    ' -- Act -------------------------------------------------
    logs(logIdx) = "3. Act: ValidarEliminacionAccion(True, " & TEST_ID_PM & ", vbYes)"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim result As Boolean
    result = ValidarEliminacionAccion(True, CStr(TEST_ID_PM), vbYes, p_Error)

    ' -- Assert ----------------------------------------------
    If p_Error <> "" Then
        logs(logIdx) = "4. Assert FAIL: p_Error populated=" & p_Error
        logIdx = logIdx + 1
        Test_ValidarEliminacionAccion_Happy = BuildFail( _
            "Unexpected error: " & p_Error, logs)
        GoTo Teardown
    End If

    If result <> True Then
        logs(logIdx) = "4. Assert FAIL: expected True, got " & CStr(result)
        logIdx = logIdx + 1
        Test_ValidarEliminacionAccion_Happy = BuildFail( _
            "Expected True (authorized + confirmed)", logs)
        GoTo Teardown
    End If

    logs(logIdx) = "4. Assert PASS: returned True"
    logIdx = logIdx + 1
    Test_ValidarEliminacionAccion_Happy = BuildOk(CStr(result), logs)
    GoTo Teardown

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.description
    logIdx = logIdx + 1
    Test_ValidarEliminacionAccion_Happy = BuildFail( _
        "Unhandled exception: " & Err.description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgosPlanMitigacionPpal WHERE IDMitigacion BETWEEN 900000 AND 999999", _
               dbFailOnError
    On Error GoTo 0
End Function

' ============================================================
' ATOM: Sad — non-existent action ID > returns False
'
' Setup:
'   p_PromptResult = vbYes (MsgBox bypass; irrelevant since
'   existence check fails first)
'   No seed rows
' Act: ValidarEliminacionAccion(False, "900099", vbYes)
' Expect: ok=true, value=False  (sad path — not an error)
' Teardown: none needed
' ============================================================
Public Function Test_ValidarEliminacionAccion_Sad() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    ' -- Arrange --------------------------------------------
    logs(logIdx) = "1. Arrange: ForceLocalBackend"
    logIdx = logIdx + 1

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ValidarEliminacionAccion_Sad = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ValidarEliminacionAccion_Sad = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' Ensure no residual rows pollute the lookup
    db.Execute "DELETE FROM TbRiesgosPlanContingenciaPpal WHERE IDContingencia BETWEEN 900000 AND 999999", _
               dbFailOnError

    logs(logIdx) = "2. Arrange: no seed rows (non-existent ID=900099)"
    logIdx = logIdx + 1

    ' -- Act -------------------------------------------------
    Const TEST_ID_MISSING As Long = 900099
    logs(logIdx) = "3. Act: ValidarEliminacionAccion(False, " & TEST_ID_MISSING & ", vbYes)"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim result As Boolean
    result = ValidarEliminacionAccion(False, CStr(TEST_ID_MISSING), vbYes, p_Error)

    ' -- Assert ----------------------------------------------
    ' Sad path — non-existent ID returns False, NOT an error
    If p_Error <> "" Then
        logs(logIdx) = "4. Assert FAIL: p_Error should be empty, got=" & p_Error
        logIdx = logIdx + 1
        Test_ValidarEliminacionAccion_Sad = BuildFail( _
            "Non-existent ID should not populate p_Error", logs)
        GoTo Teardown
    End If

    If result <> False Then
        logs(logIdx) = "4. Assert FAIL: expected False, got " & CStr(result)
        logIdx = logIdx + 1
        Test_ValidarEliminacionAccion_Sad = BuildFail( _
            "Non-existent action must return False", logs)
        GoTo Teardown
    End If

    logs(logIdx) = "4. Assert PASS: returned False for non-existent ID"
    logIdx = logIdx + 1
    Test_ValidarEliminacionAccion_Sad = BuildOk(CStr(result), logs)
    GoTo Teardown

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.description
    logIdx = logIdx + 1
    Test_ValidarEliminacionAccion_Sad = BuildFail( _
        "Unhandled exception: " & Err.description, logs)

Teardown:
    On Error Resume Next
    Set m_ObjUsuarioConectado = Nothing
    db.Execute "DELETE FROM TbRiesgosPlanMitigacionPpal WHERE IDMitigacion BETWEEN 900000 AND 999999", _
               dbFailOnError
    On Error GoTo 0
End Function

' ============================================================
' ATOM: Edge — empty string ID + whitespace-padded ID
'
' Setup: same as Sad (no seed rows needed)
' Cases tested sequentially:
'   a) p_IDAccion = ""  > expected False (empty ID)
'   b) p_IDAccion = "  900001  "  > expected False (whitespace
'      not stripped; no such row exists with padded ID)
' ============================================================
Public Function Test_ValidarEliminacionAccion_Edge() As String
    Dim logs(0 To 14) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: ForceLocalBackend"
    logIdx = logIdx + 1

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ValidarEliminacionAccion_Edge = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ValidarEliminacionAccion_Edge = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    logs(logIdx) = "2. Arrange: no seed rows"
    logIdx = logIdx + 1

    ' -- Case A: empty string --------------------------------
    Dim p_Error As String
    Dim result As Boolean

    logs(logIdx) = "3a. Act: ValidarEliminacionAccion(True, """", vbYes)"
    logIdx = logIdx + 1

    result = ValidarEliminacionAccion(True, "", vbYes, p_Error)

    If result <> False Then
        logs(logIdx) = "4a. Assert FAIL: empty ID must return False, got " & CStr(result)
        logIdx = logIdx + 1
        Test_ValidarEliminacionAccion_Edge = BuildFail( _
            "Empty ID must return False", logs)
        GoTo Teardown
    End If
    logs(logIdx) = "4a. Assert PASS: empty ID returns False"
    logIdx = logIdx + 1

    ' -- Case B: whitespace-padded ID -----------------------
    Dim paddedId As String
    paddedId = "  900001  "

    logs(logIdx) = "3b. Act: ValidarEliminacionAccion(True, """ & paddedId & """, vbYes)"
    logIdx = logIdx + 1

    p_Error = ""
    result = ValidarEliminacionAccion(True, paddedId, vbYes, p_Error)

    If result <> False Then
        logs(logIdx) = "4b. Assert FAIL: padded ID must return False, got " & CStr(result)
        logIdx = logIdx + 1
        Test_ValidarEliminacionAccion_Edge = BuildFail( _
            "Whitespace-padded ID must return False (no Trim applied)", logs)
        GoTo Teardown
    End If
    logs(logIdx) = "4b. Assert PASS: padded ID returns False"
    logIdx = logIdx + 1

    Test_ValidarEliminacionAccion_Edge = BuildOk("False+False", logs)
    GoTo Teardown

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.description
    logIdx = logIdx + 1
    Test_ValidarEliminacionAccion_Edge = BuildFail( _
        "Unhandled exception: " & Err.description, logs)

Teardown:
    On Error Resume Next
    On Error GoTo 0
End Function

' ============================================================
' ATOM: Adversarial — role mismatch (user not authorized)
'
' Setup:
'   m_ObjUsuarioConectado > un-authorized role
'   Seed: one valid PM row so the authorization is the
'         actual gate, not the existence check
' Act: ValidarEliminacionAccion(True, "900001", vbYes)
' Expect: ok=true, value=False  (authorization denied)
' Teardown: DELETE rows WHERE IDMitigacion BETWEEN 900000 AND 999999
'
' NOTE: This atom assumes the helper checks m_ObjUsuarioConectado
' internally for authorization. The user-object state is set
' via the existing m_ObjUsuarioConectado mechanism.
' ============================================================
Public Function Test_ValidarEliminacionAccion_Adversarial() As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    ' -- Arrange --------------------------------------------
    logs(logIdx) = "1. Arrange: ForceLocalBackend + set unauthorized user"
    logIdx = logIdx + 1

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        Test_ValidarEliminacionAccion_Adversarial = BuildFail("TESTS BLOCKED: " & cfgError, logs)
        Exit Function
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_ValidarEliminacionAccion_Adversarial = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' -- Seed fixture: PM row (authorization is the gate, not existence) -
    Const TEST_ID_PM As Long = 900001
    Const TEST_ID_EDICION_PM As Long = 900050
    Const TEST_ID_RIESGO_PM As Long = 900050
    Const TEST_ID_PROYECTO_PM As Long = 900060

    db.Execute "DELETE FROM TbRiesgosPlanMitigacionPpal WHERE IDMitigacion BETWEEN 900000 AND 999999", _
               dbFailOnError
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo BETWEEN 900050 AND 900059", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion BETWEEN 900050 AND 900059", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto BETWEEN 900060 AND 900069", dbFailOnError

    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, NombreProyecto) " & _
               "VALUES (" & TEST_ID_PROYECTO_PM & ", 'TEST-ADV', 'Test Adversarial')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_ID_EDICION_PM & ", " & TEST_ID_PROYECTO_PM & ", 1, 'TestAutor')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo) " & _
               "VALUES (" & TEST_ID_RIESGO_PM & ", " & TEST_ID_EDICION_PM & ", " & _
               "'PC-FIX-ADV', 'R-ADV-900050')", dbFailOnError

    db.Execute "INSERT INTO TbRiesgosPlanMitigacionPpal " & _
                "(IDMitigacion, DisparadorDelPlan, IDRiesgo) " & _
                "VALUES (" & TEST_ID_PM & ", 'PM atom adversarial', " & TEST_ID_RIESGO_PM & ")", _
               dbFailOnError
    db.Execute "INSERT INTO TbRiesgosPlanMitigacionDetalle (IDAccionMitigacion, IDMitigacion, CodAccion, Accion, ResponsableAccion, Estado) " & _
               "VALUES (" & TEST_ID_PM & ", " & TEST_ID_PM & ", 'AC-PM-ADV', 'Adversarial action', 'Test User', 'Pendiente')", _
               dbFailOnError

    logs(logIdx) = "2. Arrange: seeded PM row IDMitigacion=" & TEST_ID_PM
    logIdx = logIdx + 1

    ' -- Set unauthorized user explicitly ---------------------
    ' ForceLocalBackend short-circuits on subsequent calls
    ' (Test_Helper.bas:228) and does NOT reset m_ObjUsuarioConectado.
    ' UsuarioAutorizado (Funciones Generales.bas:799) reads that global
    ' and returns EnumSiNo.Sí if UsuarioRed is admin/calidad. To force
    ' the authorization gate, we set a non-admin Usuario here.
    Dim uNoAutorizado As New Usuario
    uNoAutorizado.UsuarioRed = "ROL_INVALIDO"
    Set m_ObjUsuarioConectado = uNoAutorizado

    ' -- Act -------------------------------------------------
    logs(logIdx) = "3. Act: ValidarEliminacionAccion(True, " & TEST_ID_PM & ", vbYes) no-auth"
    logIdx = logIdx + 1

    Dim p_Error As String
    Dim result As Boolean
    result = ValidarEliminacionAccion(True, CStr(TEST_ID_PM), vbYes, p_Error)

    ' -- Assert ----------------------------------------------
    If result <> False Then
        logs(logIdx) = "4. Assert FAIL: unauthorized user must get False, got " & CStr(result)
        logIdx = logIdx + 1
        Test_ValidarEliminacionAccion_Adversarial = BuildFail( _
            "Unauthorized user must be rejected", logs)
        GoTo Teardown
    End If

    logs(logIdx) = "4. Assert PASS: unauthorized user returns False"
    logIdx = logIdx + 1
    Test_ValidarEliminacionAccion_Adversarial = BuildOk(CStr(result), logs)
    GoTo Teardown

EH:
    logs(logIdx) = "EH: " & Err.Number & " - " & Err.description
    logIdx = logIdx + 1
    Test_ValidarEliminacionAccion_Adversarial = BuildFail( _
        "Unhandled exception: " & Err.description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgosPlanMitigacionPpal WHERE IDMitigacion BETWEEN 900000 AND 999999", _
               dbFailOnError
    On Error GoTo 0
End Function

' ============================================================
' Private helpers (same pattern as other test modules)
' ============================================================
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

