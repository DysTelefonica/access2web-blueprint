Attribute VB_Name = "Test_AvisoPublicacionHelper"
Option Compare Database
Option Explicit

' ============================================================
' Test_AvisoPublicacionHelper - TDD atoms for modAvisoPublicacionHelper
'
' Helper: AvisarPublicacion
'   Signature: Public Function AvisarPublicacion( _
'                 ByRef p_ObjEdicion As Object
'                 Optional ByRef p_PromptResult As Long
'
' Returns: 1L if row written, 0L if p_Error populated.
'
' Architecture: thin DB-only helper. The helper validates the .IDEdicion
' is not Nothing and parseable as Long, then INSERTs into TbCorreosEnviados.
' IDEdicion is typed Long in TbCorreosEnviados (post-issue-70); the helper
' uses CLng() to avoid Integer-overflow when IDEdicion > 32,767.
'
' Strategy: use real Edicion instances via Constructor.getEdicion against
' the sandbox DB (Test_Fixtures.GetTestDb). FIX_ID_* constants match
' Test_TbCorreosEnviados_IDEdicion.bas (issue-70 migration coverage) so
' we share the same fixture range and don't conflict with that test.
'
' 4 scenario classes (skill access-vba-tdd §4.5):
'   1. Happy          - IDEdicion in normal range, helper inserts 1 row
'   2. Sad            - p_ObjEdicion = Nothing, helper returns 0 + p_Error
'   3. Edge           - IDEdicion > Integer max (50000), helper persists
'                        without overflow (proves Long path works)
'   4. Adversarial    - two calls same IDEdicion, helper inserts 2 rows
'                        (no UPSERT / no dedup)
'
' Skill: access-vba-tdd v2.4.3
' SDD:    e2e-form-by-form-2026-06-22 - Bloque 2 - REQ-CAL-16A
' ============================================================

' --- Module-level constants (same numeric values as
'     Test_TbCorreosEnviados_IDEdicion.bas; redeclared locally because the
'     original module marks them Private Const). ---
Private Const FIX_ID_BASE As Long = 50000
Private Const FIX_ID_EDICION_HAPPY As Long = 50010   ' in-range (< Integer max)
Private Const FIX_ID_EDICION_EDGE As Long = 50000    ' above Integer max (32767) — edge case
Private Const FIX_ID_PROYECTO As Long = 50011
Private Const FIX_ID_EXPEDIENTE As Long = 50012

' --- JSON helpers wrapper (project convention) ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal errMsg As String, ByRef logs() As String) As String
    BuildFail = BuildJsonFail(errMsg, logs)
End Function

' --- Sandbox guard (Test_Helper public) ---
Private Function EnsureSandboxLoaded(ByRef p_Error As String) As Boolean
    EnsureSandboxLoaded = Test_Helper.EnsureTestConfigLoaded(p_Error)
End Function

' --- Get sandbox DB (Test_Fixtures public) ---
Private Function GetSandboxDb(ByRef p_Error As String) As DAO.Database
    Set GetSandboxDb = Test_Fixtures.GetTestDb(p_Error)
End Function

' -----------------------------------------------------------------------------
' Fixture helpers (FK-ordered)
'
' Mirror of SeedCorreosFixture in Test_TbCorreosEnviados_IDEdicion.bas but
' redeclared locally because that module's seed is Private Sub. We extend
' the seed to ALSO create an Edicion row at IDEdicion=50000 (the existing
' test only seeds the correo row, not the Edicion row, since it tests the
' schema level — not the application helper path).
' -----------------------------------------------------------------------------
Private Sub SeedAvisoFixture()
    On Error GoTo EH_Seed
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetSandboxDb(dbErr)
    If db Is Nothing Then Err.Raise 1001, "SeedAvisoFixture", "GetSandboxDb returned Nothing: " & dbErr

    ' Idempotent cleanup in reverse FK order — only deterministic test markers >= FIX_ID_BASE
    On Error Resume Next
    db.Execute "DELETE FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_ID_EDICION_HAPPY, dbFailOnError
    db.Execute "DELETE FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_ID_EDICION_EDGE, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_ID_EDICION_HAPPY, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_ID_EDICION_EDGE, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_ID_PROYECTO, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_ID_EXPEDIENTE, dbFailOnError
    On Error GoTo 0

    ' 1. TbExpedientes (padre)
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
               "VALUES (" & FIX_ID_EXPEDIENTE & ", 'TEST16A', 'Fixture REQ-CAL-16A', 'Test', 1)", dbFailOnError

    ' 2. TbProyectos (hijo de Expediente)
    db.Execute "INSERT INTO TbProyectos " & _
               "(IDProyecto, IDExpediente, Proyecto, ParaInformeAvisos, NombreUsuarioCalidad) " & _
               "VALUES (" & FIX_ID_PROYECTO & ", " & FIX_ID_EXPEDIENTE & ", " & _
               "'TESTPROJ16A', 'Sí', 'test_calidad_16a')", dbFailOnError

    ' 3. TbProyectosEdiciones (hijo de Proyecto) - 50010 (happy + adversarial) AND 50000 (edge)
    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDProyecto, IDEdicion, Edicion, Elaborado, FechaPreparadaParaPublicar) " & _
               "VALUES (" & FIX_ID_PROYECTO & ", " & FIX_ID_EDICION_HAPPY & ", 1, " & _
               "'test_calidad_16a', #" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDProyecto, IDEdicion, Edicion, Elaborado, FechaPreparadaParaPublicar) " & _
               "VALUES (" & FIX_ID_PROYECTO & ", " & FIX_ID_EDICION_EDGE & ", 2, " & _
               "'test_calidad_16a', #" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    ' Post-seed cardinality assertion - guard against verde-por-suerte
    Dim seedRs As DAO.Recordset
    Dim seedCountHappy As Long
    Dim seedCountEdge As Long
    Set seedRs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_ID_EDICION_HAPPY)
    If Not seedRs.EOF Then seedCountHappy = CLng(Nz(seedRs.fields("C").value, 0))
    seedRs.Close
    Set seedRs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_ID_EDICION_EDGE)
    If Not seedRs.EOF Then seedCountEdge = CLng(Nz(seedRs.fields("C").value, 0))
    seedRs.Close
    Set seedRs = Nothing
    If seedCountHappy <> 1 Or seedCountEdge <> 1 Then
        Err.Raise 1002, "SeedAvisoFixture", _
            "Post-seed assertion failed: happy=" & seedCountHappy & " (expected 1), " & _
            "edge=" & seedCountEdge & " (expected 1)"
    End If

    Set db = Nothing
    Exit Sub

EH_Seed:
    Dim eN As Long: eN = Err.Number
    Dim ed As String: ed = Err.description
    On Error Resume Next
    Set db = Nothing
    Err.Raise eN, "SeedAvisoFixture", "Seed failed: " & eN & " - " & ed
End Sub

Private Sub TeardownAvisoFixture()
    On Error Resume Next
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetSandboxDb(dbErr)
    If db Is Nothing Then Exit Sub

    db.Execute "DELETE FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_ID_EDICION_HAPPY, dbFailOnError
    db.Execute "DELETE FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_ID_EDICION_EDGE, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_ID_EDICION_HAPPY, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_ID_EDICION_EDGE, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_ID_PROYECTO, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_ID_EXPEDIENTE, dbFailOnError

    Set db = Nothing
    On Error GoTo 0
End Sub

' Build a real Edicion object via Constructor chain against the sandbox.
' Returns Nothing if Constructor fails.
Private Function BuildEdicionStub(ByVal p_IDEdicion As Long) As Edicion
    Dim m_Err As String
    Set BuildEdicionStub = Constructor.getEdicion(CStr(p_IDEdicion), m_Err)
    If m_Err <> "" Then Set BuildEdicionStub = Nothing
End Function

' ============================================================
' ATOM 1 — Happy: PublicacionExitosaRegistraUnaFilaEnTbCorreosEnviados
' GIVEN staging sandbox + Edicion at IDEdicion=50010 (in-range)
' WHEN AvisarPublicacion is called with the Edicion object
' THEN:
'   - helper returns 1
'   - p_Error empty
'   - count(TbCorreosEnviados WHERE IDEdicion=50010) == 1
'   - the row's IDEdicion column round-trips as 50010
' ============================================================
Public Function Test_AvisarPublicacion_Happy_PublicacionExitosaRegistraUnaFilaEnTbCorreosEnviados() As String
    Dim logs(0 To 9) As String
    Test_AvisarPublicacion_Happy_PublicacionExitosaRegistraUnaFilaEnTbCorreosEnviados = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded"
    logs(1) = "2. Arrange: SeedAvisoFixture (crea Edicion en 50010 + 50000)"
    logs(2) = "3. Act: BuildEdicionStub(50010) -> Edicion object"
    logs(3) = "4. Act: AvisarPublicacion(objEdicion, db, , err)"
    logs(4) = "5. Assert: helper returns 1"
    logs(5) = "6. Assert: p_Error empty"
    logs(6) = "7. Assert: count(TbCorreosEnviados WHERE IDEdicion=50010) == 1"
    logs(7) = "8. Assert: IDEdicion de la fila escrita == 50010 (round-trip)"
    logs(8) = "9. Teardown: TeardownAvisoFixture + ResetTestSession"

    Dim cfgErr As String
    If Not EnsureSandboxLoaded(cfgErr) Then
        Test_AvisarPublicacion_Happy_PublicacionExitosaRegistraUnaFilaEnTbCorreosEnviados = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    SeedAvisoFixture

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetSandboxDb(dbErr)
    If db Is Nothing Then
        Test_AvisarPublicacion_Happy_PublicacionExitosaRegistraUnaFilaEnTbCorreosEnviados = BuildFail("GetSandboxDb devolvió Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Dim m_objEdicion As Edicion
    Set m_objEdicion = BuildEdicionStub(FIX_ID_EDICION_HAPPY)
    If m_objEdicion Is Nothing Then
        Test_AvisarPublicacion_Happy_PublicacionExitosaRegistraUnaFilaEnTbCorreosEnviados = BuildFail("Constructor.getEdicion devolvió Nothing para IDEdicion=" & FIX_ID_EDICION_HAPPY, logs)
        GoTo Teardown
    End If

    Dim m_Result As Long
    Dim m_Err As String
    m_Result = AvisarPublicacion(m_objEdicion, db, , m_Err)

    If m_Err <> "" Then
        Test_AvisarPublicacion_Happy_PublicacionExitosaRegistraUnaFilaEnTbCorreosEnviados = BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If m_Result <> 1 Then
        Test_AvisarPublicacion_Happy_PublicacionExitosaRegistraUnaFilaEnTbCorreosEnviados = BuildFail("retorno esperado 1, obtuvo " & m_Result, logs)
        GoTo Teardown
    End If

    ' Cardinalidad post-INSERT (skill §4.5)
    Dim rs As DAO.Recordset
    Dim countCorreos As Long
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_ID_EDICION_HAPPY)
    If Not rs.EOF Then countCorreos = CLng(Nz(rs.fields("C").value, 0))
    rs.Close
    Set rs = Nothing
    If countCorreos <> 1 Then
        Test_AvisarPublicacion_Happy_PublicacionExitosaRegistraUnaFilaEnTbCorreosEnviados = BuildFail("expected 1 row in TbCorreosEnviados for IDEdicion=" & FIX_ID_EDICION_HAPPY & ", got " & countCorreos, logs)
        GoTo Teardown
    End If

    ' Round-trip read del IDEdicion escrito
    Dim readValue As Variant
    Set rs = db.OpenRecordset("SELECT TOP 1 IDEdicion FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_ID_EDICION_HAPPY & " ORDER BY IDCorreo DESC")
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        Test_AvisarPublicacion_Happy_PublicacionExitosaRegistraUnaFilaEnTbCorreosEnviados = BuildFail("fila insertada pero no visible en SELECT (estado inconsistente)", logs)
        GoTo Teardown
    End If
    readValue = rs.fields("IDEdicion").value
    rs.Close
    Set rs = Nothing
    If IsNull(readValue) Then
        Test_AvisarPublicacion_Happy_PublicacionExitosaRegistraUnaFilaEnTbCorreosEnviados = BuildFail("IDEdicion leío es Null", logs)
        GoTo Teardown
    End If
    If CLng(readValue) <> FIX_ID_EDICION_HAPPY Then
        Test_AvisarPublicacion_Happy_PublicacionExitosaRegistraUnaFilaEnTbCorreosEnviados = BuildFail("IDEdicion expected " & FIX_ID_EDICION_HAPPY & " but got " & CLng(readValue), logs)
        GoTo Teardown
    End If

    Test_AvisarPublicacion_Happy_PublicacionExitosaRegistraUnaFilaEnTbCorreosEnviados = BuildOk("happy_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Set m_objEdicion = Nothing
    TeardownAvisoFixture
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_AvisarPublicacion_Happy_PublicacionExitosaRegistraUnaFilaEnTbCorreosEnviados = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 2 — Sad: ObjEdicionNothing_RetornaCeroYPopulaError
' GIVEN staging sandbox
' WHEN AvisarPublicacion is called with p_ObjEdicion = Nothing
' THEN:
'   - helper returns 0
'   - p_Error is populated with "p_ObjEdicion es Nothing"
'   - no row inserted in TbCorreosEnviados (count unchanged)
' ============================================================
Public Function Test_AvisarPublicacion_Sad_ObjEdicionNothing_RetornaCeroYPopulaError() As String
    Dim logs(0 To 6) As String
    Test_AvisarPublicacion_Sad_ObjEdicionNothing_RetornaCeroYPopulaError = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded (sin seed específico)"
    logs(1) = "2. Act: AvisarPublicacion(NOTHING, db, , err)"
    logs(2) = "3. Assert: helper returns 0"
    logs(3) = "4. Assert: p_Error contains 'p_ObjEdicion es Nothing'"
    logs(4) = "5. Assert: no row inserted (count(TbCorreosEnviados) unchanged)"
    logs(5) = "6. Teardown: nothing specific + ResetTestSession"

    Dim cfgErr As String
    If Not EnsureSandboxLoaded(cfgErr) Then
        Test_AvisarPublicacion_Sad_ObjEdicionNothing_RetornaCeroYPopulaError = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetSandboxDb(dbErr)
    If db Is Nothing Then
        Test_AvisarPublicacion_Sad_ObjEdicionNothing_RetornaCeroYPopulaError = BuildFail("GetSandboxDb devolvió Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' Capture baseline count of correos to prove no row was inserted
    Dim rs As DAO.Recordset
    Dim countBefore As Long
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbCorreosEnviados")
    If Not rs.EOF Then countBefore = CLng(Nz(rs.fields("C").value, 0))
    rs.Close
    Set rs = Nothing

    ' --- Act: pass Nothing ---
    Dim m_Result As Long
    Dim m_Err As String
    m_Result = AvisarPublicacion(Nothing, db, , m_Err)

    ' --- Assert ---
    If m_Result <> 0 Then
        Test_AvisarPublicacion_Sad_ObjEdicionNothing_RetornaCeroYPopulaError = BuildFail("retorno esperado 0, obtuvo " & m_Result, logs)
        GoTo Teardown
    End If
    If Len(m_Err) = 0 Then
        Test_AvisarPublicacion_Sad_ObjEdicionNothing_RetornaCeroYPopulaError = BuildFail("p_Error esperado poblado, obtuvo vacío", logs)
        GoTo Teardown
    End If
    If InStr(1, m_Err, "p_ObjEdicion es Nothing", vbTextCompare) = 0 Then
        Test_AvisarPublicacion_Sad_ObjEdicionNothing_RetornaCeroYPopulaError = BuildFail("p_Error no contiene 'p_ObjEdicion es Nothing'. got: " & m_Err, logs)
        GoTo Teardown
    End If

    ' Cardinalidad invariante (skill §4.5): ninguna fila fue insertada
    Dim countAfter As Long
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbCorreosEnviados")
    If Not rs.EOF Then countAfter = CLng(Nz(rs.fields("C").value, 0))
    rs.Close
    Set rs = Nothing
    If countAfter <> countBefore Then
        Test_AvisarPublicacion_Sad_ObjEdicionNothing_RetornaCeroYPopulaError = BuildFail("count(TbCorreosEnviados) cambió: before=" & countBefore & " after=" & countAfter, logs)
        GoTo Teardown
    End If

    Test_AvisarPublicacion_Sad_ObjEdicionNothing_RetornaCeroYPopulaError = BuildOk("sad_nothing_rejected_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_AvisarPublicacion_Sad_ObjEdicionNothing_RetornaCeroYPopulaError = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 3 — Edge: IDEdicionSobreIntegerMax_NoDesbord
' GIVEN staging sandbox + Edicion at IDEdicion=50000 (> 32,767)
' WHEN AvisarPublicacion is called with that Edicion
' THEN:
'   - helper returns 1
'   - p_Error empty
'   - count(TbCorreosEnviados WHERE IDEdicion=50000) == 1
'   - the row's IDEdicion round-trips as exactly 50000 (no truncation, no overflow)
'
' Pre-condición documentada (issue-70): para que esta prueba sea VERDE, el
' schema de TbCorreosEnviados.IDEdicion debe ser Long (post-migration). Si
' la columna sigue siendo Integer, db.Execute fallará con DAO error 3035
' (data type conversion overflow) y el helper retornará 0 + p_Error poblado.
' Esa falla NO es bug del helper: es señal de que el sandbox necesita la
' migración (ver Test_TbCorreosEnviados_IDEdicion.bas).
' ============================================================
Public Function Test_AvisarPublicacion_Edge_IDEdicionSobreIntegerMax_NoDesbord() As String
    Dim logs(0 To 9) As String
    Test_AvisarPublicacion_Edge_IDEdicionSobreIntegerMax_NoDesbord = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded"
    logs(1) = "2. Arrange: SeedAvisoFixture (incluye Edicion en 50000)"
    logs(2) = "3. Act: BuildEdicionStub(50000) -> Edicion object (IDEdicion > Integer max)"
    logs(3) = "4. Act: AvisarPublicacion(objEdicion, db, , err)"
    logs(4) = "5. Assert: helper returns 1 (no overflow)"
    logs(5) = "6. Assert: p_Error empty"
    logs(6) = "7. Assert: count(TbCorreosEnviados WHERE IDEdicion=50000) == 1"
    logs(7) = "8. Assert: IDEdicion de la fila == 50000 (no truncado a 32767)"
    logs(8) = "9. Teardown: TeardownAvisoFixture + ResetTestSession"

    Dim cfgErr As String
    If Not EnsureSandboxLoaded(cfgErr) Then
        Test_AvisarPublicacion_Edge_IDEdicionSobreIntegerMax_NoDesbord = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    SeedAvisoFixture

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetSandboxDb(dbErr)
    If db Is Nothing Then
        Test_AvisarPublicacion_Edge_IDEdicionSobreIntegerMax_NoDesbord = BuildFail("GetSandboxDb devolvió Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Dim m_objEdicion As Edicion
    Set m_objEdicion = BuildEdicionStub(FIX_ID_EDICION_EDGE)
    If m_objEdicion Is Nothing Then
        Test_AvisarPublicacion_Edge_IDEdicionSobreIntegerMax_NoDesbord = BuildFail("Constructor.getEdicion devolvió Nothing para IDEdicion=" & FIX_ID_EDICION_EDGE, logs)
        GoTo Teardown
    End If

    Dim m_Result As Long
    Dim m_Err As String
    m_Result = AvisarPublicacion(m_objEdicion, db, , m_Err)

    If m_Err <> "" Then
        Test_AvisarPublicacion_Edge_IDEdicionSobreIntegerMax_NoDesbord = BuildFail("p_Error inesperado (revisar si sandbox tiene IDEdicion=Long post-issue-70): " & m_Err, logs)
        GoTo Teardown
    End If
    If m_Result <> 1 Then
        Test_AvisarPublicacion_Edge_IDEdicionSobreIntegerMax_NoDesbord = BuildFail("retorno esperado 1, obtuvo " & m_Result & " (sandbox sin migración Long?)", logs)
        GoTo Teardown
    End If

    ' Cardinalidad post-INSERT (skill §4.5)
    Dim rs As DAO.Recordset
    Dim countCorreos As Long
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_ID_EDICION_EDGE)
    If Not rs.EOF Then countCorreos = CLng(Nz(rs.fields("C").value, 0))
    rs.Close
    Set rs = Nothing
    If countCorreos <> 1 Then
        Test_AvisarPublicacion_Edge_IDEdicionSobreIntegerMax_NoDesbord = BuildFail("expected 1 row for IDEdicion=" & FIX_ID_EDICION_EDGE & ", got " & countCorreos, logs)
        GoTo Teardown
    End If

    ' Round-trip sin overflow
    Dim readValue As Variant
    Set rs = db.OpenRecordset("SELECT TOP 1 IDEdicion FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_ID_EDICION_EDGE & " ORDER BY IDCorreo DESC")
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        Test_AvisarPublicacion_Edge_IDEdicionSobreIntegerMax_NoDesbord = BuildFail("fila insertada pero no visible en SELECT", logs)
        GoTo Teardown
    End If
    readValue = rs.fields("IDEdicion").value
    rs.Close
    Set rs = Nothing
    If IsNull(readValue) Then
        Test_AvisarPublicacion_Edge_IDEdicionSobreIntegerMax_NoDesbord = BuildFail("IDEdicion leío es Null", logs)
        GoTo Teardown
    End If
    If CLng(readValue) <> FIX_ID_EDICION_EDGE Then
        Test_AvisarPublicacion_Edge_IDEdicionSobreIntegerMax_NoDesbord = BuildFail("IDEdicion round-trip expected " & FIX_ID_EDICION_EDGE & " but got " & CLng(readValue) & " (overflow?)", logs)
        GoTo Teardown
    End If

    Test_AvisarPublicacion_Edge_IDEdicionSobreIntegerMax_NoDesbord = BuildOk("edge_overflow_safe_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Set m_objEdicion = Nothing
    TeardownAvisoFixture
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_AvisarPublicacion_Edge_IDEdicionSobreIntegerMax_NoDesbord = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 4 — Adversarial: DobleLlamada_InsertaDosFilas
' GIVEN staging sandbox + Edicion at IDEdicion=50010
' WHEN AvisarPublicacion is called twice with the same Edicion object
' THEN:
'   - both calls return 1
'   - both p_Error empty
'   - count(TbCorreosEnviados WHERE IDEdicion=50010) == 2
'
' Decisión documentada: el helper NO hace UPSERT / dedup. Cada llamada
' inserta una fila nueva. Si en el futuro Calidad pide "una sola fila
' por edición", habrá que cambiar el helper y este átomo será
' la señal de regresión.
' ============================================================
Public Function Test_AvisarPublicacion_Adversarial_DobleLlamada_InsertaDosFilas() As String
    Dim logs(0 To 9) As String
    Test_AvisarPublicacion_Adversarial_DobleLlamada_InsertaDosFilas = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded"
    logs(1) = "2. Arrange: SeedAvisoFixture"
    logs(2) = "3. Act: BuildEdicionStub(50010)"
    logs(3) = "4. Act: AvisarPublicacion(objEdicion, db, , err)  -- 1st call"
    logs(4) = "5. Assert: 1st call returns 1, p_Error empty"
    logs(5) = "6. Act: AvisarPublicacion(objEdicion, db, , err)  -- 2nd call (same IDEdicion)"
    logs(6) = "7. Assert: 2nd call returns 1, p_Error empty"
    logs(7) = "8. Assert: count(TbCorreosEnviados WHERE IDEdicion=50010) == 2 (no dedup)"
    logs(8) = "9. Teardown: TeardownAvisoFixture + ResetTestSession"

    Dim cfgErr As String
    If Not EnsureSandboxLoaded(cfgErr) Then
        Test_AvisarPublicacion_Adversarial_DobleLlamada_InsertaDosFilas = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    SeedAvisoFixture

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetSandboxDb(dbErr)
    If db Is Nothing Then
        Test_AvisarPublicacion_Adversarial_DobleLlamada_InsertaDosFilas = BuildFail("GetSandboxDb devolvió Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Dim m_objEdicion As Edicion
    Set m_objEdicion = BuildEdicionStub(FIX_ID_EDICION_HAPPY)
    If m_objEdicion Is Nothing Then
        Test_AvisarPublicacion_Adversarial_DobleLlamada_InsertaDosFilas = BuildFail("Constructor.getEdicion devolvió Nothing para IDEdicion=" & FIX_ID_EDICION_HAPPY, logs)
        GoTo Teardown
    End If

    ' --- 1st call ---
    Dim m_Result1 As Long
    Dim m_Err1 As String
    m_Result1 = AvisarPublicacion(m_objEdicion, db, , m_Err1)
    If m_Err1 <> "" Then
        Test_AvisarPublicacion_Adversarial_DobleLlamada_InsertaDosFilas = BuildFail("1st call: p_Error no esperado: " & m_Err1, logs)
        GoTo Teardown
    End If
    If m_Result1 <> 1 Then
        Test_AvisarPublicacion_Adversarial_DobleLlamada_InsertaDosFilas = BuildFail("1st call: retorno esperado 1, obtuvo " & m_Result1, logs)
        GoTo Teardown
    End If

    ' --- 2nd call (same object, same IDEdicion) ---
    Dim m_Result2 As Long
    Dim m_Err2 As String
    m_Result2 = AvisarPublicacion(m_objEdicion, db, , m_Err2)
    If m_Err2 <> "" Then
        Test_AvisarPublicacion_Adversarial_DobleLlamada_InsertaDosFilas = BuildFail("2nd call: p_Error no esperado: " & m_Err2, logs)
        GoTo Teardown
    End If
    If m_Result2 <> 1 Then
        Test_AvisarPublicacion_Adversarial_DobleLlamada_InsertaDosFilas = BuildFail("2nd call: retorno esperado 1, obtuvo " & m_Result2, logs)
        GoTo Teardown
    End If

    ' --- Cardinalidad post-2da llamada: 2 filas (no UPSERT / no dedup) ---
    Dim rs As DAO.Recordset
    Dim countCorreos As Long
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_ID_EDICION_HAPPY)
    If Not rs.EOF Then countCorreos = CLng(Nz(rs.fields("C").value, 0))
    rs.Close
    Set rs = Nothing
    If countCorreos <> 2 Then
        Test_AvisarPublicacion_Adversarial_DobleLlamada_InsertaDosFilas = BuildFail("expected 2 rows in TbCorreosEnviados for IDEdicion=" & FIX_ID_EDICION_HAPPY & " (no dedup), got " & countCorreos, logs)
        GoTo Teardown
    End If

    Test_AvisarPublicacion_Adversarial_DobleLlamada_InsertaDosFilas = BuildOk("adversarial_double_call_two_rows_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Set m_objEdicion = Nothing
    TeardownAvisoFixture
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_AvisarPublicacion_Adversarial_DobleLlamada_InsertaDosFilas = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

