Attribute VB_Name = "Test_NotificacionPorCorreoHelper"
Option Compare Database
Option Explicit

' ============================================================
' Test_NotificacionPorCorreoHelper - TDD atoms for modNotificacionPorCorreoHelper
'
' Helper: NotificarPorCorreo
'   Signature: Public Function NotificarPorCorreo( _
'                 ByRef p_ObjRiesgo As Object
'                 Optional ByRef p_PromptResult As Long
'
' Architecture: thin pass-through. The helper validates the .Edicion.Proyecto
' chain is not Nothing, then delegates to GetBodyCorreoRiesgoMaterializado
' (Funciones Generales.bas:270) which composes the HTML body. The
' composition happens NOT to include project name/code today; this test
' module locks in the contract that the body MUST include them
' (REQ-CAL-11, bloque 2).
'
' Fixture strategy: real riesgo/Edicion/Proyecto instances via the
' Constructor chain against the sandbox DB (Test_Fixtures.GetTestDb).
' IDs are 900020..900023 to avoid collisions with the broader test fleet
' in this project.
'
' Skill: access-vba-tdd v2.4.3
' SDD:    e2e-form-by-form-2026-06-22 - Bloque 2 - REQ-CAL-11
' ============================================================

' --- JSON helpers wrapper (project convention: Test_Helper.BuildJson*) ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' --- Sandbox guard (project convention: ForceLocalBackend) ---
Private Function EnsureSandbox(ByRef logs() As String) As Boolean
    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        EnsureSandbox = False
        logs(0) = "TESTS BLOCKED: ForceLocalBackend failed: " & cfgError
    Else
        EnsureSandbox = True
    End If
End Function

' --- Helpers to build real riesgo from seeded rows ---
Private Function BuildRiesgoFromFixture(ByVal p_IDRiesgo As Long, ByRef p_db As DAO.Database) As riesgo
    Dim m_IDRiesgoStr As String
    Dim m_Err As String
    m_IDRiesgoStr = CStr(p_IDRiesgo)
    Set BuildRiesgoFromFixture = Constructor.getRiesgo(m_IDRiesgoStr, , , m_Err)
End Function

' ----------------------------------------------------------------
' ATOM 1: Happy path - body includes project name AND code
' Expected (after T-2.3):
'   - body non-empty
'   - body contains "Proyecto Alfa"  (project name from Proyecto.Proyecto)
'   - body contains "PRJ-001"       (project code from Proyecto.CodigoDocumento)
' RED until T-2.3 (body does not yet include them).
' ----------------------------------------------------------------
Public Function Test_NotificacionPorCorreo_Happy_RetornaCuerpoConNombreYcodigoProyecto() As String
    On Error GoTo EH

    Dim logs(0 To 6) As String
    logs(0) = "1. Arrange: sandbox + proyecto 'Proyecto Alfa' / 'PRJ-001' + edicion + riesgo"
    logs(1) = "2. Act: NotificarPorCorreo(m_ObjRiesgo, db, , err)"
    logs(2) = "3. Assert: body non-empty"
    logs(3) = "3a. Assert: body contains 'Proyecto Alfa'"
    logs(4) = "3b. Assert: body contains 'PRJ-001'"
    logs(5) = "4. Teardown: delete fixture rows"

    If Not EnsureSandbox(logs) Then
        Test_NotificacionPorCorreo_Happy_RetornaCuerpoConNombreYcodigoProyecto = BuildFail(logs(0), logs)
        Exit Function
    End If

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_NotificacionPorCorreo_Happy_RetornaCuerpoConNombreYcodigoProyecto = BuildFail("TESTS BLOCKED: GetTestDb devolvi\u00f3 Nothing: " & dbErr, logs)
        Exit Function
    End If

    ' --- Fixture IDs (>= 900000 per Test_Fixtures.bas FIX_ID_BASE) ---
    Const FIX_ID_EXPEDIENTE_NOTIF As Long = 900100
    Const FIX_PROJ_ID As Long = 900020
    Const FIX_EDICION_ID As Long = 900020
    Const FIX_RIESGO_ID As Long = 900020

    ' FK-ordered seed: proyecto -> edicion -> riesgo
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, CodigoDocumento, IDExpediente, Cliente) " & _
               "VALUES (" & FIX_PROJ_ID & ", 'Proyecto Alfa', 'PRJ-001', " & FIX_ID_EXPEDIENTE_NOTIF & ", 'Cliente Test')", dbFailOnError

    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado, FechaEdicion) " & _
               "VALUES (" & FIX_EDICION_ID & ", " & FIX_PROJ_ID & ", 1, 'autor-test', Date())", dbFailOnError

    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, Descripcion, DetectadoPor, FechaMaterializado, Estado, ImpactoGlobal, Calidad) " & _
               "VALUES (" & FIX_RIESGO_ID & ", " & FIX_EDICION_ID & ", 'CU-NOTIF-1', 'R-001', 'Desc test', 'Det-test', Date(), 'Materializado', 'Alto', 'Alto')", dbFailOnError

    ' --- Build real riesgo via Constructor chain ---
    Dim m_ObjRiesgo As riesgo
    Set m_ObjRiesgo = BuildRiesgoFromFixture(FIX_RIESGO_ID, db)
    If m_ObjRiesgo Is Nothing Then
        Test_NotificacionPorCorreo_Happy_RetornaCuerpoConNombreYcodigoProyecto = BuildFail("Constructor.getRiesgo devolvi\u00f3 Nothing", logs)
        GoTo Teardown
    End If

    ' --- Act ---
    Dim m_Body As String
    Dim m_Err As String
    m_Body = NotificarPorCorreo(m_ObjRiesgo, db, , m_Err)

    ' --- Assert ---
    If m_Err <> "" Then
        Test_NotificacionPorCorreo_Happy_RetornaCuerpoConNombreYcodigoProyecto = BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If

    If Len(m_Body) = 0 Then
        Test_NotificacionPorCorreo_Happy_RetornaCuerpoConNombreYcodigoProyecto = BuildFail("cuerpo vac\u00edo", logs)
        GoTo Teardown
    End If

    If InStr(1, m_Body, "Proyecto Alfa", vbTextCompare) = 0 Then
        Test_NotificacionPorCorreo_Happy_RetornaCuerpoConNombreYcodigoProyecto = BuildFail("cuerpo no contiene 'Proyecto Alfa'. body=[" & Left$(m_Body, 500) & "]", logs)
        GoTo Teardown
    End If

    If InStr(1, m_Body, "PRJ-001", vbTextCompare) = 0 Then
        Test_NotificacionPorCorreo_Happy_RetornaCuerpoConNombreYcodigoProyecto = BuildFail("cuerpo no contiene 'PRJ-001'. body=[" & Left$(m_Body, 500) & "]", logs)
        GoTo Teardown
    End If

    Test_NotificacionPorCorreo_Happy_RetornaCuerpoConNombreYcodigoProyecto = BuildOk("body_includes_name_and_code", logs)
    GoTo Teardown

EH:
    Test_NotificacionPorCorreo_Happy_RetornaCuerpoConNombreYcodigoProyecto = BuildFail("Test_NotificacionPorCorreo_Happy_RetornaCuerpoConNombreYcodigoProyecto: " & Err.description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & FIX_RIESGO_ID, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_ID, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_PROJ_ID, dbFailOnError
End Function

' ----------------------------------------------------------------
' ATOM 2: Sad path - passing Nothing returns "" with p_Error populated
' Expected (after T-2.1): validation rejects Nothing; p_Error non-empty;
'                        body returns "".
' GREEN after T-2.1 (validation lives in the helper, not in the body composition).
' ----------------------------------------------------------------
Public Function Test_NotificacionPorCorreo_Sad_ObjRiesgoNothing_RetornaVacioYPopulaError() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: sandbox only (no fixture needed - we pass Nothing)"
    logs(1) = "2. Act: NotificarPorCorreo(Nothing, db, , err)"
    logs(2) = "3. Assert: body returns empty string"
    logs(3) = "3a. Assert: p_Error non-empty"
    logs(4) = "4. Teardown: nothing to clean"

    If Not EnsureSandbox(logs) Then
        Test_NotificacionPorCorreo_Sad_ObjRiesgoNothing_RetornaVacioYPopulaError = BuildFail(logs(0), logs)
        Exit Function
    End If

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_NotificacionPorCorreo_Sad_ObjRiesgoNothing_RetornaVacioYPopulaError = BuildFail("TESTS BLOCKED: GetTestDb devolvi\u00f3 Nothing: " & dbErr, logs)
        Exit Function
    End If

    ' --- Act ---
    Dim m_Body As String
    Dim m_Err As String
    m_Body = NotificarPorCorreo(Nothing, db, , m_Err)

    ' --- Assert ---
    If m_Body <> "" Then
        Test_NotificacionPorCorreo_Sad_ObjRiesgoNothing_RetornaVacioYPopulaError = BuildFail("body esperado vac\u00edo pero se obtuvo: [" & Left$(m_Body, 200) & "]", logs)
        Exit Function
    End If

    If Len(m_Err) = 0 Then
        Test_NotificacionPorCorreo_Sad_ObjRiesgoNothing_RetornaVacioYPopulaError = BuildFail("p_Error esperado no vac\u00edo pero se obtuvo vac\u00edo", logs)
        Exit Function
    End If

    Test_NotificacionPorCorreo_Sad_ObjRiesgoNothing_RetornaVacioYPopulaError = BuildOk("nothing_rejected", logs)
    Exit Function

EH:
    Test_NotificacionPorCorreo_Sad_ObjRiesgoNothing_RetornaVacioYPopulaError = BuildFail("Test_NotificacionPorCorreo_Sad_ObjRiesgoNothing_RetornaVacioYPopulaError: " & Err.description, logs)
End Function

' ----------------------------------------------------------------
' ATOM 3: Edge - empty project name renders as 'sin nombre' placeholder
' Expected (after T-2.3): body still produced; contains 'sin nombre' as
'                          the placeholder; project code still appears.
' RED until T-2.3 (body does not yet handle empty name).
' ----------------------------------------------------------------
Public Function Test_NotificacionPorCorreo_Edge_NombreVacio_RetornaPlaceholderLegible() As String
    On Error GoTo EH

    Dim logs(0 To 6) As String
    logs(0) = "1. Arrange: sandbox + proyecto with EMPTY 'Proyecto' but valid CodigoDocumento + edicion + riesgo"
    logs(1) = "2. Act: NotificarPorCorreo(m_ObjRiesgo, db, , err)"
    logs(2) = "3. Assert: body non-empty (still composed)"
    logs(3) = "3a. Assert: body contains placeholder 'sin nombre'"
    logs(4) = "3b. Assert: body still contains project code 'PRJ-EDGE-001'"
    logs(5) = "4. Teardown: delete fixture rows"

    If Not EnsureSandbox(logs) Then
        Test_NotificacionPorCorreo_Edge_NombreVacio_RetornaPlaceholderLegible = BuildFail(logs(0), logs)
        Exit Function
    End If

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_NotificacionPorCorreo_Edge_NombreVacio_RetornaPlaceholderLegible = BuildFail("TESTS BLOCKED: GetTestDb devolvi\u00f3 Nothing: " & dbErr, logs)
        Exit Function
    End If

    ' --- Fixture IDs ---
    Const FIX_PROJ_ID As Long = 900021
    Const FIX_EDICION_ID As Long = 900021
    Const FIX_RIESGO_ID As Long = 900021

    ' Proyecto.Proyecto = '' (empty string - edge case)
    Const FIX_ID_EXPEDIENTE_EDGE As Long = 900101
    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, CodigoDocumento, IDExpediente, Cliente) " & _
               "VALUES (" & FIX_PROJ_ID & ", '', 'PRJ-EDGE-001', " & FIX_ID_EXPEDIENTE_EDGE & ", 'Cliente Edge')", dbFailOnError

    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado, FechaEdicion) " & _
               "VALUES (" & FIX_EDICION_ID & ", " & FIX_PROJ_ID & ", 1, 'autor-edge', Date())", dbFailOnError

    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, Descripcion, DetectadoPor, FechaMaterializado, Estado, ImpactoGlobal, Calidad) " & _
               "VALUES (" & FIX_RIESGO_ID & ", " & FIX_EDICION_ID & ", 'CU-NOTIF-EDGE', 'R-EDGE', 'Desc edge', 'Det-edge', Date(), 'Materializado', 'Medio', 'Medio')", dbFailOnError

    ' --- Build real riesgo ---
    Dim m_ObjRiesgo As riesgo
    Set m_ObjRiesgo = BuildRiesgoFromFixture(FIX_RIESGO_ID, db)
    If m_ObjRiesgo Is Nothing Then
        Test_NotificacionPorCorreo_Edge_NombreVacio_RetornaPlaceholderLegible = BuildFail("Constructor.getRiesgo devolvi\u00f3 Nothing", logs)
        GoTo Teardown
    End If

    ' --- Act ---
    Dim m_Body As String
    Dim m_Err As String
    m_Body = NotificarPorCorreo(m_ObjRiesgo, db, , m_Err)

    ' --- Assert ---
    If m_Err <> "" Then
        Test_NotificacionPorCorreo_Edge_NombreVacio_RetornaPlaceholderLegible = BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If

    If Len(m_Body) = 0 Then
        Test_NotificacionPorCorreo_Edge_NombreVacio_RetornaPlaceholderLegible = BuildFail("cuerpo vac\u00edo (deber\u00eda componerse con placeholder)", logs)
        GoTo Teardown
    End If

    If InStr(1, m_Body, "sin nombre", vbTextCompare) = 0 Then
        Test_NotificacionPorCorreo_Edge_NombreVacio_RetornaPlaceholderLegible = BuildFail("cuerpo no contiene placeholder 'sin nombre'. body=[" & Left$(m_Body, 500) & "]", logs)
        GoTo Teardown
    End If

    If InStr(1, m_Body, "PRJ-EDGE-001", vbTextCompare) = 0 Then
        Test_NotificacionPorCorreo_Edge_NombreVacio_RetornaPlaceholderLegible = BuildFail("cuerpo no contiene 'PRJ-EDGE-001' (c\u00f3digo s\u00ed deber\u00eda aparecer). body=[" & Left$(m_Body, 500) & "]", logs)
        GoTo Teardown
    End If

    Test_NotificacionPorCorreo_Edge_NombreVacio_RetornaPlaceholderLegible = BuildOk("empty_name_placeholder", logs)
    GoTo Teardown

EH:
    Test_NotificacionPorCorreo_Edge_NombreVacio_RetornaPlaceholderLegible = BuildFail("Test_NotificacionPorCorreo_Edge_NombreVacio_RetornaPlaceholderLegible: " & Err.description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & FIX_RIESGO_ID, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_ID, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_PROJ_ID, dbFailOnError
End Function

' ----------------------------------------------------------------
' ATOM 4: Adversarial - db=Nothing must not raise
' Expected (after T-2.1): helper accepts db=Nothing (composition is pure);
'                        returns a valid body without error.
' GREEN after T-2.1 (helper doesn't actually read db).
' ----------------------------------------------------------------
Public Function Test_NotificacionPorCorreo_Adversarial_DbNothing_NoFalla() As String
    On Error GoTo EH

    Dim logs(0 To 6) As String
    logs(0) = "1. Arrange: sandbox + proyecto + edicion + riesgo"
    logs(1) = "2. Act: NotificarPorCorreo(m_ObjRiesgo, NOTHING, , err)  <- adversarial db"
    logs(2) = "3. Assert: no exception raised (we reach this line)"
    logs(3) = "3a. Assert: body still produced (helper is pure - db unused)"
    logs(4) = "3b. Assert: p_Error empty (success path)"
    logs(5) = "4. Teardown: delete fixture rows"

    If Not EnsureSandbox(logs) Then
        Test_NotificacionPorCorreo_Adversarial_DbNothing_NoFalla = BuildFail(logs(0), logs)
        Exit Function
    End If

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_NotificacionPorCorreo_Adversarial_DbNothing_NoFalla = BuildFail("TESTS BLOCKED: GetTestDb devolvi\u00f3 Nothing: " & dbErr, logs)
        Exit Function
    End If

    ' --- Fixture IDs ---
    Const FIX_ID_EXPEDIENTE_ADV As Long = 900102
    Const FIX_PROJ_ID As Long = 900022
    Const FIX_EDICION_ID As Long = 900022
    Const FIX_RIESGO_ID As Long = 900022

    db.Execute "INSERT INTO TbProyectos (IDProyecto, Proyecto, CodigoDocumento, IDExpediente, Cliente) " & _
               "VALUES (" & FIX_PROJ_ID & ", 'Proyecto Adversarial', 'PRJ-ADV-001', " & FIX_ID_EXPEDIENTE_ADV & ", 'Cliente Adv')", dbFailOnError

    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado, FechaEdicion) " & _
               "VALUES (" & FIX_EDICION_ID & ", " & FIX_PROJ_ID & ", 1, 'autor-adv', Date())", dbFailOnError

    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, Descripcion, DetectadoPor, FechaMaterializado, Estado, ImpactoGlobal, Calidad) " & _
               "VALUES (" & FIX_RIESGO_ID & ", " & FIX_EDICION_ID & ", 'CU-NOTIF-ADV', 'R-ADV', 'Desc adv', 'Det-adv', Date(), 'Materializado', 'Bajo', 'Bajo')", dbFailOnError

    ' --- Build real riesgo (we still need DB to populate IDEdicion etc.) ---
    Dim m_ObjRiesgo As riesgo
    Set m_ObjRiesgo = BuildRiesgoFromFixture(FIX_RIESGO_ID, db)
    If m_ObjRiesgo Is Nothing Then
        Test_NotificacionPorCorreo_Adversarial_DbNothing_NoFalla = BuildFail("Constructor.getRiesgo devolvi\u00f3 Nothing", logs)
        GoTo Teardown
    End If

    ' --- Act: pass db = NOTHING (the adversarial case) ---
    Dim m_Body As String
    Dim m_Err As String
    m_Body = NotificarPorCorreo(m_ObjRiesgo, Nothing, , m_Err)

    ' --- Assert: helper accepted db=Nothing and composed the body ---
    If m_Err <> "" Then
        Test_NotificacionPorCorreo_Adversarial_DbNothing_NoFalla = BuildFail("p_Error inesperado con db=Nothing: " & m_Err, logs)
        GoTo Teardown
    End If

    If Len(m_Body) = 0 Then
        Test_NotificacionPorCorreo_Adversarial_DbNothing_NoFalla = BuildFail("cuerpo vac\u00edo con db=Nothing (helper deber\u00eda ser puro)", logs)
        GoTo Teardown
    End If

    Test_NotificacionPorCorreo_Adversarial_DbNothing_NoFalla = BuildOk("db_nothing_no_crash", logs)
    GoTo Teardown

EH:
    Test_NotificacionPorCorreo_Adversarial_DbNothing_NoFalla = BuildFail("Test_NotificacionPorCorreo_Adversarial_DbNothing_NoFalla: " & Err.description, logs)

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & FIX_RIESGO_ID, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_ID, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_PROJ_ID, dbFailOnError
End Function

