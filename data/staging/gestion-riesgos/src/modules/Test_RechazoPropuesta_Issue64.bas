Attribute VB_Name = "Test_RechazoPropuesta_Issue64"
Option Compare Database
Option Explicit

' ============================================================
' Test_RechazoPropuesta_Issue64 — Tests para el fix de issue #64
'
' Skill: access-vba-tdd v2.5 (schema-first + fixture patterns + HandleError)
'
' Cobertura (6 tests atómicos, REQ-EML-064-01..05):
'   Test 1 — InsertaFilaEnTbCorreosEnviados (happy path, ParaInformeAvisos="Sí")
'   Test 2 — RegistraTrazaEnTbProyectoEdicionesCorreoRevision (happy path)
'   Test 3 — ParaInformeAvisosNo_NoEnviaCorreo_NoMuestraOK (silent no-op, nothing-guard)
'   Test 4 — MotivoVacio_NoEnvia_NoRegistra (pre-condition preserved)
'   Test 5 — EsRechazoPropuestaNotificado_NothingGuard (helper unit test)
'   Test 6 — HTMLSafeEnCuerpo_MotivoConCaracteresRaros (XSS prevention)
'
' Fixture IDs: rango FIX_ID_BASE = 30000 (deterministic, safe range — see lines 22-26 for
' why this is below 32767: TbCorreosEnviados.IDEdicion and TbUsuariosAplicaciones.Id
' are Integer columns, max 32767. Range 30000-30999 is safely above real data and
' within Integer bounds. Follow-up issue #70 to ALTER COLUMN LONG.)
' ============================================================

' --- Module-level declarations (AGENTS.md rule 3) ---
' NOTE: Fixture IDs MUST be < 32,767 because TbCorreosEnviados.IDEdicion and
' TbUsuariosAplicaciones.Id are Integer (type:3, size:2). Real data uses
' small IDs (< 10000). Test range 30000-30999 is safely above real data
' and within Integer bounds.
Private Const FIX_ID_BASE As Long = 30000
Private Const FIX_EXPEDIENTE_ID As Long = 30010
Private Const FIX_PROYECTO_ID As Long = 30011
Private Const FIX_EDICION_ID As Long = 30012
Private Const FIX_USUARIO_ID As Long = 30013
Private Const FIX_USUARIO_CALIDAD_RED As String = "test_calidad_64"
Private Const FIX_MOTIVO As String = "Falta evidencia test_64"

' --- JSON helpers (delegación a Test_Helper) ---
Private Function JsonOk(ByVal value As Variant, ByRef logs() As String) As String
    JsonOk = BuildJsonOk(value, logs)
End Function

Private Function JsonFail(ByVal errMsg As String, ByRef logs() As String) As String
    JsonFail = BuildJsonFail(errMsg, logs)
End Function

' --- EnsureTestConfigLoaded (delegación a Test_Helper) ---
Private Function EnsureTestConfigLoaded(ByRef p_Error As String) As Boolean
    EnsureTestConfigLoaded = Test_Helper.EnsureTestConfigLoaded(p_Error)
End Function

' --- GetTestDb (delegación a Test_Fixtures) ---
Private Function GetTestDb(ByRef p_Error As String) As DAO.Database
    Set GetTestDb = Test_Fixtures.GetTestDb(p_Error)
End Function

' --- Fixture helpers (FK-ordered) ---

' Seed fixture for ParaInformeAvisos="Sí" scenario (tests 1, 2, 5, 6)
Private Sub SeedRechazoFixture(Optional ByVal p_ParaInformeAvisos As String = "Sí")
    On Error GoTo EH_Seed
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then Err.Raise 1001, "SeedRechazoFixture", "GetTestDb returned Nothing: " & dbErr

    ' Idempotent cleanup in reverse FK order
    On Error Resume Next
    db.Execute "DELETE FROM TbProyectoEdicionesCorreoRevision WHERE IDEdicion=" & FIX_EDICION_ID, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_ID, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_PROYECTO_ID, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_EXPEDIENTE_ID, dbFailOnError
    db.Execute "DELETE FROM TbUsuariosAplicaciones WHERE Id=" & FIX_USUARIO_ID, dbFailOnError
    On Error GoTo 0

    ' 1. TbExpedientes (padre)
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
               "VALUES (" & FIX_EXPEDIENTE_ID & ", 'TEST64', 'Fixture issue 64', 'Test', 1)", dbFailOnError

    ' 2. TbUsuariosAplicaciones (quality user)
    db.Execute "INSERT INTO TbUsuariosAplicaciones " & _
               "(Id, CorreoUsuario, UsuarioRed, Nombre, Activado) " & _
               "VALUES (" & FIX_USUARIO_ID & ", 'test_calidad_64@test.local', " & _
               "'" & FIX_USUARIO_CALIDAD_RED & "', 'Test Calidad 64', True)", dbFailOnError

    ' 3. TbProyectos (hijo de Expediente)
    db.Execute "INSERT INTO TbProyectos " & _
               "(IDProyecto, IDExpediente, Proyecto, ParaInformeAvisos, NombreUsuarioCalidad) " & _
               "VALUES (" & FIX_PROYECTO_ID & ", " & FIX_EXPEDIENTE_ID & ", " & _
               "'TESTPROJ64', '" & p_ParaInformeAvisos & "', '" & FIX_USUARIO_CALIDAD_RED & "')", dbFailOnError

    ' 4. TbProyectosEdiciones (hijo de Proyecto)
    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDEdicion, IDProyecto, Edicion, Elaborado, FechaPreparadaParaPublicar) " & _
               "VALUES (" & FIX_EDICION_ID & ", " & FIX_PROYECTO_ID & ", 1, " & _
               "'" & FIX_USUARIO_CALIDAD_RED & "', #" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    Set db = Nothing
    Exit Sub

EH_Seed:
    Dim e As Long

    e = Err.Number
    Dim d As String

    d = Err.Description
    On Error Resume Next
    Set db = Nothing
    Err.Raise e, "SeedRechazoFixture", "Seed failed: " & e & " - " & d
End Sub

' Teardown in reverse FK order — only deletes deterministic test markers (>= FIX_ID_BASE)
Private Sub TeardownRechazoFixture()
    On Error Resume Next
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then Exit Sub

    db.Execute "DELETE FROM TbProyectoEdicionesCorreoRevision WHERE IDEdicion=" & FIX_EDICION_ID, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_ID, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_PROYECTO_ID, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_EXPEDIENTE_ID, dbFailOnError
    db.Execute "DELETE FROM TbUsuariosAplicaciones WHERE Id=" & FIX_USUARIO_ID, dbFailOnError
    db.Execute "DELETE FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_EDICION_ID, dbFailOnError

    Set db = Nothing
    On Error GoTo 0
End Sub

' Load Edicion via Constructor so Proyecto + lazy refs are populated
Private Function LoadTestEdicion(ByRef p_Error As String) As Edicion
    On Error GoTo EH
    Dim ed As Edicion
    Set ed = Constructor.getEdicion(CStr(FIX_EDICION_ID), p_Error)
    If p_Error <> "" Then Set ed = Nothing
    Set LoadTestEdicion = ed
    Exit Function
EH:
    p_Error = "LoadTestEdicion: " & Err.Description
    Set LoadTestEdicion = Nothing
End Function

' ============================================================
' RunAll — Aggregator following Test_PublicabilidadEdicion_RunAll pattern
' ============================================================
Public Function Test_RechazoPropuesta_Issue64_RunAll() As String
    Dim results(0 To 5) As String
    Dim names(0 To 5) As String
    Dim i As Long
    Dim runError As String
    Dim outLogs(0 To 9) As String

    ' SuiteSetup
    ForceLocalBackend runError
    If runError <> "" Then
        outLogs(0) = "TESTS BLOCKED: " & runError
        Test_RechazoPropuesta_Issue64_RunAll = JsonFail("TESTS BLOCKED: " & runError, outLogs)
        Exit Function
    End If
    outLogs(0) = "SuiteSetup OK"

    names(0) = "Test_RechazoPropuesta_Issue64_InsertaFilaEnTbCorreosEnviados"
    names(1) = "Test_RechazoPropuesta_Issue64_RegistraTrazaEnTbProyectoEdicionesCorreoRevision"
    names(2) = "Test_RechazoPropuesta_Issue64_ParaInformeAvisosNo_NoEnviaCorreo_NoMuestraOK"
    names(3) = "Test_RechazoPropuesta_Issue64_MotivoVacio_NoEnvia_NoRegistra"
    names(4) = "Test_RechazoPropuesta_Issue64_EsRechazoPropuestaNotificado_NothingGuard"
    names(5) = "Test_RechazoPropuesta_Issue64_HTMLSafeEnCuerpo_MotivoConCaracteresRaros"

    results(0) = Test_RechazoPropuesta_Issue64_InsertaFilaEnTbCorreosEnviados()
    results(1) = Test_RechazoPropuesta_Issue64_RegistraTrazaEnTbProyectoEdicionesCorreoRevision()
    results(2) = Test_RechazoPropuesta_Issue64_ParaInformeAvisosNo_NoEnviaCorreo_NoMuestraOK()
    results(3) = Test_RechazoPropuesta_Issue64_MotivoVacio_NoEnvia_NoRegistra()
    results(4) = Test_RechazoPropuesta_Issue64_EsRechazoPropuestaNotificado_NothingGuard()
    results(5) = Test_RechazoPropuesta_Issue64_HTMLSafeEnCuerpo_MotivoConCaracteresRaros()

    ' Teardown defensivo
    TeardownRechazoFixture
    outLogs(1) = "Teardown OK"

    ' SuiteTeardown
    Test_Helper.ResetTestSession
    outLogs(2) = "SuiteTeardown OK"

    ' Acumular
    Dim allOk As Boolean
    allOk = True
    Dim firstFailure As String
    firstFailure = ""
    For i = 0 To 5
        If InStr(results(i), """ok"":false") > 0 Then
            allOk = False
            If firstFailure = "" Then firstFailure = names(i)
        End If
        outLogs(i + 3) = names(i) & ": " & IIf(InStr(results(i), """ok"":false") > 0, "FAIL", "OK")
    Next i

    If allOk Then
        Test_RechazoPropuesta_Issue64_RunAll = JsonOk("all_pass", outLogs)
    Else
        Test_RechazoPropuesta_Issue64_RunAll = JsonFail("some_tests_failed: " & firstFailure, outLogs)
    End If
End Function

' ============================================================
' Test 1 — InsertaFilaEnTbCorreosEnviados (happy path, REQ-EML-064-01)
' GIVEN ParaInformeAvisos="Sí", motivo non-empty
' WHEN RechazoPropuestaParaPublicacion runs
' THEN count(TbCorreosEnviados WHERE IDEdicion=E.IDEdicion) = 1
'   AND Cuerpo LIKE "*Falta evidencia*"
'   AND IDEdicion column populated (not blank)
' ============================================================
Public Function Test_RechazoPropuesta_Issue64_InsertaFilaEnTbCorreosEnviados() As String
    On Error GoTo EH
    Dim logs(0 To 9) As String
    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedRechazoFixture(ParaInformeAvisos=Sí)"
    logs(2) = "3. Arrange: LoadTestEdicion + motivo"
    logs(3) = "4. Act: RechazoPropuestaParaPublicacion"
    logs(4) = "5. Assert: count(TbCorreosEnviados) == 1"
    logs(5) = "6. Assert: Cuerpo LIKE *motivo*"
    logs(6) = "7. Assert: IDEdicion column populated"
    logs(7) = "8. Teardown: TeardownRechazoFixture"
    logs(8) = "9. Teardown: Test_Helper.ResetTestSession"
    logs(9) = "10. Teardown: drop extra correo rows from fixture"

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_RechazoPropuesta_Issue64_InsertaFilaEnTbCorreosEnviados = JsonFail(cfgError, logs)
        Exit Function
    End If

    SeedRechazoFixture p_ParaInformeAvisos:="Sí"

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RechazoPropuesta_Issue64_InsertaFilaEnTbCorreosEnviados = JsonFail("GetTestDb returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Dim ed As Edicion
    Dim pError As String
    Set ed = LoadTestEdicion(pError)
    If ed Is Nothing Then
        Test_RechazoPropuesta_Issue64_InsertaFilaEnTbCorreosEnviados = JsonFail("LoadTestEdicion returned Nothing: " & pError, logs)
        GoTo Teardown
    End If

    ' Post-seed assertion: prove the seed actually persisted.
    Dim seedRs1 As DAO.Recordset
    Set seedRs1 = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_ID)
    Dim seedCount1 As Long
    If Not seedRs1.EOF Then seedCount1 = CLng(Nz(seedRs1.Fields("C").value, 0))
    seedRs1.Close
    Set seedRs1 = Nothing
    If seedCount1 <> 1 Then
        Test_RechazoPropuesta_Issue64_InsertaFilaEnTbCorreosEnviados = JsonFail("Post-seed assertion failed: expected 1 row in TbProyectosEdiciones, got " & seedCount1, logs)
        GoTo Teardown
    End If

    ed.PropuestaRechazadaPorCalidadMotivo = FIX_MOTIVO

    Dim correoRechazo As CORREO
    Set correoRechazo = ed.RechazoPropuestaParaPublicacion(pError)

    Dim capturedError As String
    capturedError = pError

    Dim countCorreos As Long
    Dim cuerpoValue As String
    Dim ideValue As Variant
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_EDICION_ID)
    If Not rs.EOF Then countCorreos = CLng(Nz(rs.Fields("C").value, 0))
    rs.Close
    Set rs = db.OpenRecordset("SELECT TOP 1 Cuerpo, IDEdicion FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_EDICION_ID & " ORDER BY IDCorreo DESC")
    If Not rs.EOF Then
        cuerpoValue = CStr(Nz(rs.Fields("Cuerpo").value, ""))
        ideValue = rs.Fields("IDEdicion").value
    End If
    rs.Close
    Set rs = Nothing

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownRechazoFixture
    Test_Helper.ResetTestSession
    On Error GoTo 0

    If Len(capturedError) > 0 Then
        Test_RechazoPropuesta_Issue64_InsertaFilaEnTbCorreosEnviados = JsonFail("pError from RechazoPropuestaParaPublicacion: " & capturedError, logs)
        Exit Function
    End If
    If countCorreos <> 1 Then
        Test_RechazoPropuesta_Issue64_InsertaFilaEnTbCorreosEnviados = JsonFail("Expected 1 row in TbCorreosEnviados but got " & countCorreos, logs)
        Exit Function
    End If
    If InStr(1, cuerpoValue, FIX_MOTIVO, vbTextCompare) = 0 Then
        Test_RechazoPropuesta_Issue64_InsertaFilaEnTbCorreosEnviados = JsonFail("Cuerpo does not contain motivo '" & FIX_MOTIVO & "'. Got: " & Left$(cuerpoValue, 200), logs)
        Exit Function
    End If
    If IsNull(ideValue) Then
        Test_RechazoPropuesta_Issue64_InsertaFilaEnTbCorreosEnviados = JsonFail("IDEdicion column is Null (expected populated)", logs)
        Exit Function
    End If
    If CLng(ideValue) <> FIX_EDICION_ID Then
        Test_RechazoPropuesta_Issue64_InsertaFilaEnTbCorreosEnviados = JsonFail("IDEdicion expected " & FIX_EDICION_ID & " but got " & CLng(ideValue), logs)
        Exit Function
    End If

    Test_RechazoPropuesta_Issue64_InsertaFilaEnTbCorreosEnviados = JsonOk("inserta_fila_tb_correos_enviados_pass", logs)
    Exit Function

EH:
    Test_RechazoPropuesta_Issue64_InsertaFilaEnTbCorreosEnviados = JsonFail("unexpected: " & Err.Description, logs)
    Resume Teardown
End Function

' ============================================================
' Test 2 — RegistraTrazaEnTbProyectoEdicionesCorreoRevision (REQ-EML-064-01)
' GIVEN same setup as Test 1
' WHEN RechazoPropuestaParaPublicacion runs
' THEN count(TbProyectoEdicionesCorreoRevision WHERE IDEdicion=E.IDEdicion) = 1
'   AND IDCorreo matches the new TbCorreosEnviados.IDCorreo
' ============================================================
Public Function Test_RechazoPropuesta_Issue64_RegistraTrazaEnTbProyectoEdicionesCorreoRevision() As String
    On Error GoTo EH
    Dim logs(0 To 8) As String
    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedRechazoFixture(ParaInformeAvisos=Sí)"
    logs(2) = "3. Arrange: LoadTestEdicion + motivo"
    logs(3) = "4. Act: RechazoPropuestaParaPublicacion"
    logs(4) = "5. Assert: count(TbProyectoEdicionesCorreoRevision) == 1"
    logs(5) = "6. Assert: IDCorreo matches TbCorreosEnviados.IDCorreo"
    logs(6) = "7. Teardown: TeardownRechazoFixture"
    logs(7) = "8. Teardown: Test_Helper.ResetTestSession"
    logs(8) = "9. Teardown: drop extra correo rows"

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_RechazoPropuesta_Issue64_RegistraTrazaEnTbProyectoEdicionesCorreoRevision = JsonFail(cfgError, logs)
        Exit Function
    End If

    SeedRechazoFixture p_ParaInformeAvisos:="Sí"

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RechazoPropuesta_Issue64_RegistraTrazaEnTbProyectoEdicionesCorreoRevision = JsonFail("GetTestDb returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Dim ed As Edicion
    Dim pError As String
    Set ed = LoadTestEdicion(pError)
    If ed Is Nothing Then
        Test_RechazoPropuesta_Issue64_RegistraTrazaEnTbProyectoEdicionesCorreoRevision = JsonFail("LoadTestEdicion returned Nothing: " & pError, logs)
        GoTo Teardown
    End If

    ' Post-seed assertion: prove the seed actually persisted.
    Dim seedRs2 As DAO.Recordset
    Set seedRs2 = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_ID)
    Dim seedCount2 As Long
    If Not seedRs2.EOF Then seedCount2 = CLng(Nz(seedRs2.Fields("C").value, 0))
    seedRs2.Close
    Set seedRs2 = Nothing
    If seedCount2 <> 1 Then
        Test_RechazoPropuesta_Issue64_RegistraTrazaEnTbProyectoEdicionesCorreoRevision = JsonFail("Post-seed assertion failed: expected 1 row in TbProyectosEdiciones, got " & seedCount2, logs)
        GoTo Teardown
    End If

    ed.PropuestaRechazadaPorCalidadMotivo = FIX_MOTIVO

    Dim correoRechazo As CORREO
    Set correoRechazo = ed.RechazoPropuestaParaPublicacion(pError)

    Dim countRevision As Long
    Dim revisionIDCorreo As Variant
    Dim correoID As Variant
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbProyectoEdicionesCorreoRevision WHERE IDEdicion=" & FIX_EDICION_ID)
    If Not rs.EOF Then countRevision = CLng(Nz(rs.Fields("C").value, 0))
    rs.Close
    Set rs = db.OpenRecordset("SELECT TOP 1 IDCorreo FROM TbProyectoEdicionesCorreoRevision WHERE IDEdicion=" & FIX_EDICION_ID & " ORDER BY IDEnvioCorreoTecnico DESC")
    If Not rs.EOF Then revisionIDCorreo = rs.Fields("IDCorreo").value
    rs.Close
    Set rs = db.OpenRecordset("SELECT TOP 1 IDCorreo FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_EDICION_ID & " ORDER BY IDCorreo DESC")
    If Not rs.EOF Then correoID = rs.Fields("IDCorreo").value
    rs.Close
    Set rs = Nothing

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownRechazoFixture
    Test_Helper.ResetTestSession
    On Error GoTo 0

    If countRevision <> 1 Then
        Test_RechazoPropuesta_Issue64_RegistraTrazaEnTbProyectoEdicionesCorreoRevision = JsonFail("Expected 1 row in TbProyectoEdicionesCorreoRevision but got " & countRevision, logs)
        Exit Function
    End If
    If IsNull(revisionIDCorreo) Or IsNull(correoID) Then
        Test_RechazoPropuesta_Issue64_RegistraTrazaEnTbProyectoEdicionesCorreoRevision = JsonFail("IDCorreo is Null in revision or correos", logs)
        Exit Function
    End If
    If CLng(revisionIDCorreo) <> CLng(correoID) Then
        Test_RechazoPropuesta_Issue64_RegistraTrazaEnTbProyectoEdicionesCorreoRevision = JsonFail("revision.IDCorreo=" & CLng(revisionIDCorreo) & " but correos.IDCorreo=" & CLng(correoID), logs)
        Exit Function
    End If

    Test_RechazoPropuesta_Issue64_RegistraTrazaEnTbProyectoEdicionesCorreoRevision = JsonOk("registra_traza_tb_proyecto_ediciones_correo_revision_pass", logs)
    Exit Function

EH:
    Test_RechazoPropuesta_Issue64_RegistraTrazaEnTbProyectoEdicionesCorreoRevision = JsonFail("unexpected: " & Err.Description, logs)
    Resume Teardown
End Function

' ============================================================
' Test 3 — ParaInformeAvisosNo_NoEnviaCorreo_NoMuestraOK (REQ-EML-064-02, REQ-EML-064-04)
' GIVEN ParaInformeAvisos="No"
' WHEN RechazoPropuestaParaPublicacion runs
' THEN count(TbCorreosEnviados WHERE IDEdicion=E.IDEdicion) unchanged
'   AND count(TbProyectoEdicionesCorreoRevision WHERE IDEdicion=E.IDEdicion) unchanged
'   AND PropuestaRechazadaPorCalidadMotivo persisted in TbProyectosEdiciones
'   AND EsRechazoPropuestaNotificado(m_CorreoRechazo) = False
' ============================================================
Public Function Test_RechazoPropuesta_Issue64_ParaInformeAvisosNo_NoEnviaCorreo_NoMuestraOK() As String
    On Error GoTo EH
    Dim logs(0 To 9) As String
    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedRechazoFixture(ParaInformeAvisos=No)"
    logs(2) = "3. Arrange: LoadTestEdicion"
    logs(3) = "4. Act: RechazoPropuestaParaPublicacion"
    logs(4) = "5. Assert: count(TbCorreosEnviados) == 0"
    logs(5) = "6. Assert: count(TbProyectoEdicionesCorreoRevision) == 0"
    logs(6) = "7. Assert: PropuestaRechazadaPorCalidadMotivo persisted"
    logs(7) = "8. Assert: EsRechazoPropuestaNotificado(m_CorreoRechazo) = False"
    logs(8) = "9. Teardown: TeardownRechazoFixture"
    logs(9) = "10. Teardown: Test_Helper.ResetTestSession"

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_RechazoPropuesta_Issue64_ParaInformeAvisosNo_NoEnviaCorreo_NoMuestraOK = JsonFail(cfgError, logs)
        Exit Function
    End If

    SeedRechazoFixture p_ParaInformeAvisos:="No"

    Dim ed As Edicion
    Dim pError As String
    Set ed = LoadTestEdicion(pError)
    If ed Is Nothing Then
        Test_RechazoPropuesta_Issue64_ParaInformeAvisosNo_NoEnviaCorreo_NoMuestraOK = JsonFail("LoadTestEdicion returned Nothing: " & pError, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetTestDb(dbErr)
    Dim countCorreos As Long
    Dim countRevision As Long
    Dim motivoPersistido As String
    If db Is Nothing Then
        Test_RechazoPropuesta_Issue64_ParaInformeAvisosNo_NoEnviaCorreo_NoMuestraOK = JsonFail("GetTestDb returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' Post-seed assertion: prove the seed actually persisted.
    Dim seedRs3 As DAO.Recordset
    Set seedRs3 = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_ID)
    Dim seedCount3 As Long
    If Not seedRs3.EOF Then seedCount3 = CLng(Nz(seedRs3.Fields("C").value, 0))
    seedRs3.Close
    Set seedRs3 = Nothing
    If seedCount3 <> 1 Then
        Test_RechazoPropuesta_Issue64_ParaInformeAvisosNo_NoEnviaCorreo_NoMuestraOK = JsonFail("Post-seed assertion failed: expected 1 row in TbProyectosEdiciones, got " & seedCount3, logs)
        GoTo Teardown
    End If

    ed.PropuestaRechazadaPorCalidadMotivo = FIX_MOTIVO

    Dim correoRechazo As CORREO
    Set correoRechazo = ed.RechazoPropuestaParaPublicacion(pError)

    ' Capture before teardown
    Dim capturedError As String
    capturedError = pError
    Dim correoIsNothing As Boolean
    correoIsNothing = (correoRechazo Is Nothing)
    Dim correoID As String
    correoID = ""
    If Not correoRechazo Is Nothing Then
        correoID = correoRechazo.IDCorreo
    End If
    Dim notificadoResult As Boolean
    notificadoResult = EsRechazoPropuestaNotificado(correoRechazo)

    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_EDICION_ID)
    If Not rs.EOF Then countCorreos = CLng(Nz(rs.Fields("C").value, 0))
    rs.Close
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbProyectoEdicionesCorreoRevision WHERE IDEdicion=" & FIX_EDICION_ID)
    If Not rs.EOF Then countRevision = CLng(Nz(rs.Fields("C").value, 0))
    rs.Close
    Set rs = db.OpenRecordset("SELECT PropuestaRechazadaPorCalidadMotivo FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_ID)
    If Not rs.EOF Then motivoPersistido = CStr(Nz(rs.Fields("PropuestaRechazadaPorCalidadMotivo").value, ""))
    rs.Close
    Set rs = Nothing

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownRechazoFixture
    Test_Helper.ResetTestSession
    On Error GoTo 0

    If countCorreos <> 0 Then
        Test_RechazoPropuesta_Issue64_ParaInformeAvisosNo_NoEnviaCorreo_NoMuestraOK = JsonFail("Expected 0 rows in TbCorreosEnviados but got " & countCorreos, logs)
        Exit Function
    End If
    If countRevision <> 0 Then
        Test_RechazoPropuesta_Issue64_ParaInformeAvisosNo_NoEnviaCorreo_NoMuestraOK = JsonFail("Expected 0 rows in TbProyectoEdicionesCorreoRevision but got " & countRevision, logs)
        Exit Function
    End If
    If motivoPersistido <> FIX_MOTIVO Then
        Test_RechazoPropuesta_Issue64_ParaInformeAvisosNo_NoEnviaCorreo_NoMuestraOK = JsonFail("Expected PropuestaRechazadaPorCalidadMotivo='" & FIX_MOTIVO & "' but got '" & motivoPersistido & "'", logs)
        Exit Function
    End If
    If notificadoResult <> False Then
        Test_RechazoPropuesta_Issue64_ParaInformeAvisosNo_NoEnviaCorreo_NoMuestraOK = JsonFail("Expected EsRechazoPropuestaNotificado=False but got True (IDCorreo='" & correoID & "')", logs)
        Exit Function
    End If

    Test_RechazoPropuesta_Issue64_ParaInformeAvisosNo_NoEnviaCorreo_NoMuestraOK = JsonOk("para_informe_avisos_no_no_envia_pass", logs)
    Exit Function

EH:
    Test_RechazoPropuesta_Issue64_ParaInformeAvisosNo_NoEnviaCorreo_NoMuestraOK = JsonFail("unexpected: " & Err.Description, logs)
    Resume Teardown
End Function

' ============================================================
' Test 4 — MotivoVacio_NoEnvia_NoRegistra (REQ-EML-064-03)
' GIVEN motivo empty
' WHEN RechazoPropuestaParaPublicacion runs
' THEN p_Error set
'   AND count(TbCorreosEnviados) unchanged
'   AND count(TbProyectoEdicionesCorreoRevision) unchanged
'   AND PropuestaRechazadaPorCalidadMotivo unchanged (rejection not persisted)
' ============================================================
Public Function Test_RechazoPropuesta_Issue64_MotivoVacio_NoEnvia_NoRegistra() As String
    On Error GoTo EH
    Dim logs(0 To 9) As String
    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedRechazoFixture(ParaInformeAvisos=Sí)"
    logs(2) = "3. Arrange: LoadTestEdicion + PropuestaRechazadaPorCalidadMotivo=''"
    logs(3) = "4. Act: RechazoPropuestaParaPublicacion"
    logs(4) = "5. Assert: p_Error set (motivo vacío guard)"
    logs(5) = "6. Assert: count(TbCorreosEnviados) == 0"
    logs(6) = "7. Assert: count(TbProyectoEdicionesCorreoRevision) == 0"
    logs(7) = "8. Assert: PropuestaRechazadaPorCalidadMotivo unchanged (empty)"
    logs(8) = "9. Teardown: TeardownRechazoFixture"
    logs(9) = "10. Teardown: Test_Helper.ResetTestSession"

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_RechazoPropuesta_Issue64_MotivoVacio_NoEnvia_NoRegistra = JsonFail(cfgError, logs)
        Exit Function
    End If

    SeedRechazoFixture p_ParaInformeAvisos:="Sí"

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RechazoPropuesta_Issue64_MotivoVacio_NoEnvia_NoRegistra = JsonFail("GetTestDb returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Dim ed As Edicion
    Dim pError As String
    Set ed = LoadTestEdicion(pError)
    If ed Is Nothing Then
        Test_RechazoPropuesta_Issue64_MotivoVacio_NoEnvia_NoRegistra = JsonFail("LoadTestEdicion returned Nothing: " & pError, logs)
        GoTo Teardown
    End If

    ' Post-seed assertion: prove the seed actually persisted.
    Dim seedRs4 As DAO.Recordset
    Set seedRs4 = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_ID)
    Dim seedCount4 As Long
    If Not seedRs4.EOF Then seedCount4 = CLng(Nz(seedRs4.Fields("C").value, 0))
    seedRs4.Close
    Set seedRs4 = Nothing
    If seedCount4 <> 1 Then
        Test_RechazoPropuesta_Issue64_MotivoVacio_NoEnvia_NoRegistra = JsonFail("Post-seed assertion failed: expected 1 row in TbProyectosEdiciones, got " & seedCount4, logs)
        GoTo Teardown
    End If

    ed.PropuestaRechazadaPorCalidadMotivo = ""

    Dim correoRechazo As CORREO
    Set correoRechazo = ed.RechazoPropuestaParaPublicacion(pError)

    Dim capturedError As String
    capturedError = pError

    Dim countCorreos As Long
    Dim countRevision As Long
    Dim motivoPersistido As String
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_EDICION_ID)
    If Not rs.EOF Then countCorreos = CLng(Nz(rs.Fields("C").value, 0))
    rs.Close
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbProyectoEdicionesCorreoRevision WHERE IDEdicion=" & FIX_EDICION_ID)
    If Not rs.EOF Then countRevision = CLng(Nz(rs.Fields("C").value, 0))
    rs.Close
    Set rs = db.OpenRecordset("SELECT PropuestaRechazadaPorCalidadMotivo FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_ID)
    If Not rs.EOF Then motivoPersistido = CStr(Nz(rs.Fields("PropuestaRechazadaPorCalidadMotivo").value, ""))
    rs.Close
    Set rs = Nothing

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownRechazoFixture
    Test_Helper.ResetTestSession
    On Error GoTo 0

    If capturedError = "" Then
        Test_RechazoPropuesta_Issue64_MotivoVacio_NoEnvia_NoRegistra = JsonFail("Expected p_Error set for empty motivo but got empty", logs)
        Exit Function
    End If
    If countCorreos <> 0 Then
        Test_RechazoPropuesta_Issue64_MotivoVacio_NoEnvia_NoRegistra = JsonFail("Expected 0 rows in TbCorreosEnviados but got " & countCorreos, logs)
        Exit Function
    End If
    If countRevision <> 0 Then
        Test_RechazoPropuesta_Issue64_MotivoVacio_NoEnvia_NoRegistra = JsonFail("Expected 0 rows in TbProyectoEdicionesCorreoRevision but got " & countRevision, logs)
        Exit Function
    End If
    If motivoPersistido <> "" Then
        Test_RechazoPropuesta_Issue64_MotivoVacio_NoEnvia_NoRegistra = JsonFail("Expected PropuestaRechazadaPorCalidadMotivo empty (unchanged) but got '" & motivoPersistido & "'", logs)
        Exit Function
    End If

    Test_RechazoPropuesta_Issue64_MotivoVacio_NoEnvia_NoRegistra = JsonOk("motivo_vacio_no_envia_pass", logs)
    Exit Function

EH:
    Test_RechazoPropuesta_Issue64_MotivoVacio_NoEnvia_NoRegistra = JsonFail("unexpected: " & Err.Description, logs)
    Resume Teardown
End Function

' ============================================================
' Test 5 — EsRechazoPropuestaNotificado_NothingGuard (REQ-EML-064-04)
' Direct unit test of the helper:
'   - Nothing → False (no error 91)
'   - IDCorreo = "" → False
'   - IDCorreo = "C1" → True
' Verifies split-statement pattern (no IIf, no And/Or with Is Nothing on LHS).
' ============================================================
Public Function Test_RechazoPropuesta_Issue64_EsRechazoPropuestaNotificado_NothingGuard() As String
    On Error GoTo EH
    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: nothing correo"
    logs(1) = "2. Act: EsRechazoPropuestaNotificado(Nothing)"
    logs(2) = "3. Assert: returns False without error 91"
    logs(3) = "4. Assert: IDCorreo='' returns False, IDCorreo='C1' returns True"
    logs(4) = "5. Teardown: n/a (pure unit test)"

    ' Case 1: Nothing -> False
    Dim resultNothing As Boolean
    resultNothing = EsRechazoPropuestaNotificado(Nothing)
    If resultNothing <> False Then
        Test_RechazoPropuesta_Issue64_EsRechazoPropuestaNotificado_NothingGuard = JsonFail("Expected False for Nothing but got True", logs)
        Exit Function
    End If

    ' Case 2: IDCorreo = "" -> False
    Dim correoVacio As CORREO
    Set correoVacio = New CORREO
    correoVacio.IDCorreo = ""
    Dim resultVacio As Boolean
    resultVacio = EsRechazoPropuestaNotificado(correoVacio)
    Set correoVacio = Nothing
    If resultVacio <> False Then
        Test_RechazoPropuesta_Issue64_EsRechazoPropuestaNotificado_NothingGuard = JsonFail("Expected False for IDCorreo='' but got True", logs)
        Exit Function
    End If

    ' Case 3: IDCorreo = "C1" -> True
    Dim correoConID As CORREO
    Set correoConID = New CORREO
    correoConID.IDCorreo = "C1"
    Dim resultConID As Boolean
    resultConID = EsRechazoPropuestaNotificado(correoConID)
    Set correoConID = Nothing
    If resultConID <> True Then
        Test_RechazoPropuesta_Issue64_EsRechazoPropuestaNotificado_NothingGuard = JsonFail("Expected True for IDCorreo='C1' but got False", logs)
        Exit Function
    End If

    Test_RechazoPropuesta_Issue64_EsRechazoPropuestaNotificado_NothingGuard = JsonOk("nothing_guard_pass", logs)
    Exit Function

EH:
    Test_RechazoPropuesta_Issue64_EsRechazoPropuestaNotificado_NothingGuard = JsonFail("unexpected: " & Err.Description & " (source: " & Err.Source & ")", logs)
End Function

' ============================================================
' Test 6 — HTMLSafeEnCuerpo_MotivoConCaracteresRaros (REQ-EML-064-05)
' GIVEN motivo = "<script>x</script> & ""q"""
' WHEN EnviarCorreoRechazoPropuestaPublicacion runs
' THEN TbCorreosEnviados.Cuerpo contains &lt;script&gt;, &amp;, &quot;
'   AND does NOT contain raw "<script"
' ============================================================
Public Function Test_RechazoPropuesta_Issue64_HTMLSafeEnCuerpo_MotivoConCaracteresRaros() As String
    On Error GoTo EH
    Dim logs(0 To 8) As String
    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedRechazoFixture(ParaInformeAvisos=Sí)"
    logs(2) = "3. Arrange: LoadTestEdicion + XSS motivo"
    logs(3) = "4. Act: RechazoPropuestaParaPublicacion"
    logs(4) = "5. Assert: Cuerpo contains &lt;script&gt;, &amp;, &quot;"
    logs(5) = "6. Assert: Cuerpo does NOT contain raw <script"
    logs(6) = "7. Teardown: TeardownRechazoFixture"
    logs(7) = "8. Teardown: Test_Helper.ResetTestSession"
    logs(8) = "9. Teardown: drop extra correo rows"

    Const MOTIVO_XSS As String = "<script>x</script> & ""q"""

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_RechazoPropuesta_Issue64_HTMLSafeEnCuerpo_MotivoConCaracteresRaros = JsonFail(cfgError, logs)
        Exit Function
    End If

    SeedRechazoFixture p_ParaInformeAvisos:="Sí"

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RechazoPropuesta_Issue64_HTMLSafeEnCuerpo_MotivoConCaracteresRaros = JsonFail("GetTestDb returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Dim ed As Edicion
    Dim pError As String
    Set ed = LoadTestEdicion(pError)
    If ed Is Nothing Then
        Test_RechazoPropuesta_Issue64_HTMLSafeEnCuerpo_MotivoConCaracteresRaros = JsonFail("LoadTestEdicion returned Nothing: " & pError, logs)
        GoTo Teardown
    End If

    ' Post-seed assertion: prove the seed actually persisted.
    Dim seedRs6 As DAO.Recordset
    Set seedRs6 = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_ID)
    Dim seedCount6 As Long
    If Not seedRs6.EOF Then seedCount6 = CLng(Nz(seedRs6.Fields("C").value, 0))
    seedRs6.Close
    Set seedRs6 = Nothing
    If seedCount6 <> 1 Then
        Test_RechazoPropuesta_Issue64_HTMLSafeEnCuerpo_MotivoConCaracteresRaros = JsonFail("Post-seed assertion failed: expected 1 row in TbProyectosEdiciones, got " & seedCount6, logs)
        GoTo Teardown
    End If

    ed.PropuestaRechazadaPorCalidadMotivo = MOTIVO_XSS

    Dim correoRechazo As CORREO
    Set correoRechazo = ed.RechazoPropuestaParaPublicacion(pError)

    Dim cuerpoValue As String
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT TOP 1 Cuerpo FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_EDICION_ID & " ORDER BY IDCorreo DESC")
    If Not rs.EOF Then cuerpoValue = CStr(Nz(rs.Fields("Cuerpo").value, ""))
    rs.Close
    Set rs = Nothing

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownRechazoFixture
    Test_Helper.ResetTestSession
    On Error GoTo 0

    If Len(cuerpoValue) = 0 Then
        Test_RechazoPropuesta_Issue64_HTMLSafeEnCuerpo_MotivoConCaracteresRaros = JsonFail("No TbCorreosEnviados row found for IDEdicion=" & FIX_EDICION_ID, logs)
        Exit Function
    End If
    If InStr(1, cuerpoValue, "&lt;script&gt;", vbBinaryCompare) = 0 Then
        Test_RechazoPropuesta_Issue64_HTMLSafeEnCuerpo_MotivoConCaracteresRaros = JsonFail("Cuerpo does not contain '&lt;script&gt;'. Got: " & Left$(cuerpoValue, 300), logs)
        Exit Function
    End If
    If InStr(1, cuerpoValue, "&amp;", vbBinaryCompare) = 0 Then
        Test_RechazoPropuesta_Issue64_HTMLSafeEnCuerpo_MotivoConCaracteresRaros = JsonFail("Cuerpo does not contain '&amp;'. Got: " & Left$(cuerpoValue, 300), logs)
        Exit Function
    End If
    If InStr(1, cuerpoValue, "&quot;", vbBinaryCompare) = 0 Then
        Test_RechazoPropuesta_Issue64_HTMLSafeEnCuerpo_MotivoConCaracteresRaros = JsonFail("Cuerpo does not contain '&quot;'. Got: " & Left$(cuerpoValue, 300), logs)
        Exit Function
    End If
    If InStr(1, cuerpoValue, "<script>", vbBinaryCompare) > 0 Then
        Test_RechazoPropuesta_Issue64_HTMLSafeEnCuerpo_MotivoConCaracteresRaros = JsonFail("Cuerpo contains RAW '<script>' (XSS!). Got: " & Left$(cuerpoValue, 300), logs)
        Exit Function
    End If

    Test_RechazoPropuesta_Issue64_HTMLSafeEnCuerpo_MotivoConCaracteresRaros = JsonOk("html_safe_en_cuerpo_pass", logs)
    Exit Function

EH:
    Test_RechazoPropuesta_Issue64_HTMLSafeEnCuerpo_MotivoConCaracteresRaros = JsonFail("unexpected: " & Err.Description, logs)
    Resume Teardown
End Function
