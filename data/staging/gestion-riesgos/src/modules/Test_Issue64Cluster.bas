Attribute VB_Name = "Test_Issue64Cluster"
Option Compare Database
Option Explicit

' ============================================================
' Test_Issue64Cluster — Atoms para los fixes de issues #64 cluster
'
'   #66 (root cause) — Correo.EnviarCorreo: silent no-op observable via p_Error
'   #67 (related)    — CuerpoHTMLConEstiloCorporativo: <meta charset="UTF-8">
'   #68 (defense)    — RechazoPropuestaParaPublicacion: guard EsCalidad
'
' Skill: access-vba-tdd v2.6 (schema-first + fixture patterns + HandleError)
'
' Cobertura (3 átomos atómicos, REQ-EML-066-01, REQ-EML-067-01, REQ-EML-068-01):
'   Atom 1 — Test_Issue64Cluster_EnviarCorreo_NotificaNoOpCuandoParaInformeAvisosEsNo
'   Atom 2 — Test_Issue64Cluster_CuerpoHTMLConEstiloCorporativo_DeclaraCharsetUTF8
'   Atom 3 — Test_Issue64Cluster_RechazoPropuestaParaPublicacion_RechazaSinPermisosCalidad
'
' Fixture IDs: rango 31000-31999 (separado de Test_RechazoPropuesta_Issue64 que usa
' 30000-30999). Idempotency via DELETE filtrado por IDEdicion/IDProyecto.
' ============================================================

' --- Module-level constants (FIX_*, all distinct from Test_RechazoPropuesta_Issue64) ---
Private Const FIX_ID_BASE As Long = 31000
Private Const FIX_EXPEDIENTE_ID As Long = 31010
Private Const FIX_PROYECTO_NO_AVISOS_ID As Long = 31011
Private Const FIX_PROYECTO_SI_AVISOS_ID As Long = 31012
Private Const FIX_EDICION_NO_AVISOS_ID As Long = 31013
Private Const FIX_EDICION_SI_AVISOS_ID As Long = 31014
Private Const FIX_USUARIO_ID As Long = 31015
Private Const FIX_USUARIO_RED As String = "test_issue64cluster_68"

' --- JSON helpers (delegación a Test_Helper) ---
Private Function JsonOk(ByVal value As Variant, ByRef logs() As String) As String
    JsonOk = BuildJsonOk(value, logs)
End Function

Private Function JsonFail(ByVal errMsg As String, ByRef logs() As String) As String
    JsonFail = BuildJsonFail(errMsg, logs)
End Function

' --- Role state save/restore (para Atom 3) ---
Private Sub SaveCurrentRoleState(ByRef p_EsAdministrador As EnumSiNo, ByRef p_EsCalidad As EnumSiNo, ByRef p_EsTecnico As EnumSiNo)
    p_EsAdministrador = EsAdministrador
    p_EsCalidad = EsCalidad
    p_EsTecnico = EsTecnico
End Sub

Private Sub RestoreCurrentRoleState(ByVal p_EsAdministrador As EnumSiNo, ByRef p_EsCalidad As EnumSiNo, ByRef p_EsTecnico As EnumSiNo)
    EsAdministrador = p_EsAdministrador
    EsCalidad = p_EsCalidad
    EsTecnico = p_EsTecnico
End Sub

' --- Idempotent seed: FIX_EXPEDIENTE_ID + FIX_USUARIO_ID + 2 proyectos + 2 ediciones ---
' Atom 1 usa ParaInformeAvisos=No, Atom 3 usa ParaInformeAvisos=Si.
Private Sub SeedAll_Cluster(Optional ByVal cfgError As String = "")
    On Error GoTo EH_Seed
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then err.Raise 1001, "SeedAll_Cluster", "GetTestDb returned Nothing: " & dbErr

    ' Idempotent cleanup in reverse FK order
    On Error Resume Next
    db.Execute "DELETE FROM TbProyectoEdicionesCorreoRevision WHERE IDEdicion IN (" & FIX_EDICION_NO_AVISOS_ID & "," & FIX_EDICION_SI_AVISOS_ID & ")", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion IN (" & FIX_EDICION_NO_AVISOS_ID & "," & FIX_EDICION_SI_AVISOS_ID & ")", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto IN (" & FIX_PROYECTO_NO_AVISOS_ID & "," & FIX_PROYECTO_SI_AVISOS_ID & ")", dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_EXPEDIENTE_ID, dbFailOnError
    db.Execute "DELETE FROM TbUsuariosAplicaciones WHERE Id=" & FIX_USUARIO_ID, dbFailOnError
    On Error GoTo 0

    ' 1. TbExpedientes (padre)
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
               "VALUES (" & FIX_EXPEDIENTE_ID & ", 'TEST64C', 'Fixture issue 64 cluster', 'Test', 1)", dbFailOnError

    ' 2. TbUsuariosAplicaciones
    db.Execute "INSERT INTO TbUsuariosAplicaciones " & _
               "(Id, CorreoUsuario, UsuarioRed, Nombre, Activado) " & _
               "VALUES (" & FIX_USUARIO_ID & ", 'test_64c@test.local', " & _
               "'" & FIX_USUARIO_RED & "', 'Test 64 Cluster', True)", dbFailOnError

    ' 3. TbProyectos: 1 con ParaInformeAvisos=No (para Atom 1) + 1 con Si (para Atom 3)
    db.Execute "INSERT INTO TbProyectos " & _
               "(IDProyecto, IDExpediente, Proyecto, ParaInformeAvisos, NombreUsuarioCalidad) " & _
               "VALUES (" & FIX_PROYECTO_NO_AVISOS_ID & ", " & FIX_EXPEDIENTE_ID & ", " & _
               "'TESTPROJ64C_NO', 'No', '" & FIX_USUARIO_RED & "')", dbFailOnError
    db.Execute "INSERT INTO TbProyectos " & _
               "(IDProyecto, IDExpediente, Proyecto, ParaInformeAvisos, NombreUsuarioCalidad) " & _
               "VALUES (" & FIX_PROYECTO_SI_AVISOS_ID & ", " & FIX_EXPEDIENTE_ID & ", " & _
               "'TESTPROJ64C_SI', 'Sí', '" & FIX_USUARIO_RED & "')", dbFailOnError

    ' 4. TbProyectosEdiciones (1 por proyecto)
    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDEdicion, IDProyecto, Edicion, Elaborado, FechaPreparadaParaPublicar) " & _
               "VALUES (" & FIX_EDICION_NO_AVISOS_ID & ", " & FIX_PROYECTO_NO_AVISOS_ID & ", 1, " & _
               "'" & FIX_USUARIO_RED & "', #" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones " & _
               "(IDEdicion, IDProyecto, Edicion, Elaborado, FechaPreparadaParaPublicar) " & _
               "VALUES (" & FIX_EDICION_SI_AVISOS_ID & ", " & FIX_PROYECTO_SI_AVISOS_ID & ", 1, " & _
               "'" & FIX_USUARIO_RED & "', #" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

    Set db = Nothing
    Exit Sub

EH_Seed:
    Dim errNum As Long
    errNum = err.Number
    Dim eDesc As String
    eDesc = err.description
    On Error Resume Next
    Set db = Nothing
    err.Raise errNum, "SeedAll_Cluster", "Seed failed: " & errNum & " - " & eDesc
End Sub

' --- Teardown: borra todas las filas de fixture ---
Private Sub TeardownAll_Cluster()
    On Error Resume Next
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then Exit Sub
    db.Execute "DELETE FROM TbProyectoEdicionesCorreoRevision WHERE IDEdicion IN (" & FIX_EDICION_NO_AVISOS_ID & "," & FIX_EDICION_SI_AVISOS_ID & ")", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion IN (" & FIX_EDICION_NO_AVISOS_ID & "," & FIX_EDICION_SI_AVISOS_ID & ")", dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto IN (" & FIX_PROYECTO_NO_AVISOS_ID & "," & FIX_PROYECTO_SI_AVISOS_ID & ")", dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_EXPEDIENTE_ID, dbFailOnError
    db.Execute "DELETE FROM TbUsuariosAplicaciones WHERE Id=" & FIX_USUARIO_ID, dbFailOnError
    db.Execute "DELETE FROM TbCorreosEnviados WHERE IDEdicion IN (" & FIX_EDICION_NO_AVISOS_ID & "," & FIX_EDICION_SI_AVISOS_ID & ")", dbFailOnError
    Set db = Nothing
    On Error GoTo 0
End Sub

' ============================================================
' RunAll — Aggregator
' ============================================================
Public Function Test_Issue64Cluster_RunAll() As String
    Dim results(0 To 2) As String
    Dim names(0 To 2) As String
    Dim i As Long
    Dim outLogs(0 To 6) As String

    Dim runError As String
    ForceLocalBackend runError
    If runError <> "" Then
        outLogs(0) = "TESTS BLOCKED: " & runError
        Test_Issue64Cluster_RunAll = JsonFail("TESTS BLOCKED: " & runError, outLogs)
        Exit Function
    End If
    outLogs(0) = "SuiteSetup OK"

    names(0) = "Test_Issue64Cluster_EnviarCorreo_NotificaNoOpCuandoParaInformeAvisosEsNo"
    names(1) = "Test_Issue64Cluster_CuerpoHTMLConEstiloCorporativo_DeclaraCharsetUTF8"
    names(2) = "Test_Issue64Cluster_RechazoPropuestaParaPublicacion_RechazaSinPermisosCalidad"

    results(0) = Test_Issue64Cluster_EnviarCorreo_NotificaNoOpCuandoParaInformeAvisosEsNo()
    results(1) = Test_Issue64Cluster_CuerpoHTMLConEstiloCorporativo_DeclaraCharsetUTF8()
    results(2) = Test_Issue64Cluster_RechazoPropuestaParaPublicacion_RechazaSinPermisosCalidad()

    TeardownAll_Cluster
    outLogs(1) = "Teardown OK"
    Test_Helper.ResetTestSession
    outLogs(2) = "SuiteTeardown OK"

    Dim allOk As Boolean
    allOk = True
    Dim firstFailure As String
    firstFailure = ""
    For i = 0 To 2
        If InStr(results(i), """ok"":false") > 0 Then
            allOk = False
            If firstFailure = "" Then firstFailure = names(i)
        End If
        outLogs(i + 3) = names(i) & ": " & IIf(InStr(results(i), """ok"":false") > 0, "FAIL", "OK")
    Next i

    If allOk Then
        Test_Issue64Cluster_RunAll = JsonOk("all_pass", outLogs)
    Else
        Test_Issue64Cluster_RunAll = JsonFail("some_tests_failed: " & firstFailure, outLogs)
    End If
End Function

' ============================================================
' Atom 1 — EnviarCorreo_NotificaNoOpCuandoParaInformeAvisosEsNo (REQ-EML-066-01)
' GIVEN correo with Edicion.Proyecto.ParaInformeAvisos="No"
' WHEN EnviarCorreo runs
' THEN p_Error is non-empty (sentinel observable para el caller)
'   AND p_Error contains the no-op descriptor
'   AND correo.IDCorreo stays empty (no se asignó ID)
'   AND no row in TbCorreosEnviados
' ============================================================
Public Function Test_Issue64Cluster_EnviarCorreo_NotificaNoOpCuandoParaInformeAvisosEsNo() As String
    On Error GoTo EH
    Dim logs(0 To 7) As String
    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedAll_Cluster (FIX_EDICION_NO_AVISOS_ID with ParaInformeAvisos=No)"
    logs(2) = "3. Arrange: LoadTestEdicion + build correo"
    logs(3) = "4. Act: correo.EnviarCorreo pError"
    logs(4) = "5. Assert: pError is non-empty"
    logs(5) = "6. Assert: correo.IDCorreo stays empty"
    logs(6) = "7. Assert: count(TbCorreosEnviados) == 0"
    logs(7) = "8. Teardown: TeardownAll_Cluster"

    Dim cfgError As String
    If Not Test_Helper.EnsureTestConfigLoaded(cfgError) Then
        Test_Issue64Cluster_EnviarCorreo_NotificaNoOpCuandoParaInformeAvisosEsNo = JsonFail(cfgError, logs)
        Exit Function
    End If
    SeedAll_Cluster

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Issue64Cluster_EnviarCorreo_NotificaNoOpCuandoParaInformeAvisosEsNo = JsonFail("GetTestDb returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    Dim ed As Edicion
    Dim loadErr As String
    Set ed = Constructor.getEdicion(CStr(FIX_EDICION_NO_AVISOS_ID), loadErr)
    If ed Is Nothing Then
        Test_Issue64Cluster_EnviarCorreo_NotificaNoOpCuandoParaInformeAvisosEsNo = JsonFail("Constructor.getEdicion failed: " & loadErr, logs)
        GoTo Teardown
    End If

    ' Build correo
    Dim c As correo
    Set c = New correo
    c.IDEdicion = CStr(FIX_EDICION_NO_AVISOS_ID)
    Set c.Edicion = ed
    c.Asunto = "Test issue 66 no-op"
    c.Cuerpo = "Cuerpo de prueba"
    c.Destinatarios = "test_dest@example.com"
    c.FechaGrabacion = Now()

    ' Act
    Dim pError As String
    c.EnviarCorreo pError

    ' Read DB state
    Dim countCorreos As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_EDICION_NO_AVISOS_ID)
    If Not rs.EOF Then countCorreos = CLng(Nz(rs.fields("C").value, 0))
    rs.Close
    Set rs = Nothing

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownAll_Cluster
    Test_Helper.ResetTestSession
    On Error GoTo 0

    ' Assertions
    If Len(pError) = 0 Then
        Test_Issue64Cluster_EnviarCorreo_NotificaNoOpCuandoParaInformeAvisosEsNo = JsonFail("pError must be non-empty (silent no-op not fixed). pError='" & pError & "'", logs)
        Exit Function
    End If
    ' The sentinel must mention no-op semantics (callers depend on it being detectable)
    If InStr(1, pError, "no se envia correo", vbTextCompare) = 0 And InStr(1, pError, "no se envía correo", vbTextCompare) = 0 And InStr(1, pError, "no-op", vbTextCompare) = 0 And InStr(1, pError, "no_op", vbTextCompare) = 0 And InStr(1, pError, "ParaInformeAvisos", vbTextCompare) = 0 Then
        Test_Issue64Cluster_EnviarCorreo_NotificaNoOpCuandoParaInformeAvisosEsNo = JsonFail("pError must signal no-op (e.g. 'no se envia correo' or 'ParaInformeAvisos'). Got: '" & pError & "'", logs)
        Exit Function
    End If
    If Len(c.IDCorreo) > 0 Then
        Test_Issue64Cluster_EnviarCorreo_NotificaNoOpCuandoParaInformeAvisosEsNo = JsonFail("correo.IDCorreo must stay empty on no-op. Got: '" & c.IDCorreo & "'", logs)
        Exit Function
    End If
    If countCorreos <> 0 Then
        Test_Issue64Cluster_EnviarCorreo_NotificaNoOpCuandoParaInformeAvisosEsNo = JsonFail("Expected 0 rows in TbCorreosEnviados but got " & countCorreos, logs)
        Exit Function
    End If

    Test_Issue64Cluster_EnviarCorreo_NotificaNoOpCuandoParaInformeAvisosEsNo = JsonOk("no_op_observable_via_p_error", logs)
    Exit Function

EH:
    Test_Issue64Cluster_EnviarCorreo_NotificaNoOpCuandoParaInformeAvisosEsNo = JsonFail("unexpected: " & err.description, logs)
    Resume Teardown
End Function

' ============================================================
' Atom 2 — CuerpoHTMLConEstiloCorporativo_DeclaraCharsetUTF8 (REQ-EML-067-01)
' GIVEN plain text body (no embedded <html>)
' WHEN CuerpoHTMLConEstiloCorporativo runs
' THEN output contains <meta charset="UTF-8"> inside <head>
'   AND non-ASCII subject chars are preserved verbatim in <title>
' ============================================================
Public Function Test_Issue64Cluster_CuerpoHTMLConEstiloCorporativo_DeclaraCharsetUTF8() As String
    On Error GoTo EH
    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: subject with non-ASCII (áéíóúñ), plain body"
    logs(1) = "2. Act: CuerpoHTMLConEstiloCorporativo"
    logs(2) = "3. Assert: <head> precedes <meta charset=UTF-8>"
    logs(3) = "4. Assert: charset=UTF-8 token is present"
    logs(4) = "5. Assert: non-ASCII subject chars preserved in output"

    Const SUBJECT_UTF8 As String = "Motivación: áéíóú — Propuesta rechazada"
    Const BODY_PLAIN As String = "Motivo: Falta evidencia en el plan de mitigación"

    Dim html As String
    html = CuerpoHTMLConEstiloCorporativo(SUBJECT_UTF8, BODY_PLAIN)

    ' Find <head> position
    Dim headPos As Long
    headPos = InStr(1, html, "<head>", vbTextCompare)
    If headPos = 0 Then
        Test_Issue64Cluster_CuerpoHTMLConEstiloCorporativo_DeclaraCharsetUTF8 = JsonFail("Output has no <head> tag. Got: " & Left$(html, 300), logs)
        Exit Function
    End If

    ' Find <meta charset="UTF-8"> position (accept single or double quotes)
    Dim charsetPos As Long
    charsetPos = InStr(1, html, "charset=""UTF-8""", vbBinaryCompare)
    If charsetPos = 0 Then
        charsetPos = InStr(1, html, "charset='UTF-8'", vbBinaryCompare)
    End If
    If charsetPos = 0 Then
        Test_Issue64Cluster_CuerpoHTMLConEstiloCorporativo_DeclaraCharsetUTF8 = JsonFail("Output has no <meta charset='UTF-8'>. Got head: " & Mid$(html, headPos, 200), logs)
        Exit Function
    End If

    ' Assert charset is inside <head> (charsetPos must come after headPos, and before </head>)
    If charsetPos < headPos Then
        Test_Issue64Cluster_CuerpoHTMLConEstiloCorporativo_DeclaraCharsetUTF8 = JsonFail("charset meta is before <head>. headPos=" & headPos & " charsetPos=" & charsetPos, logs)
        Exit Function
    End If
    Dim headEndPos As Long
    headEndPos = InStr(headPos, html, "</head>", vbTextCompare)
    If headEndPos > 0 And charsetPos > headEndPos Then
        Test_Issue64Cluster_CuerpoHTMLConEstiloCorporativo_DeclaraCharsetUTF8 = JsonFail("charset meta is after </head>. charsetPos=" & charsetPos & " headEndPos=" & headEndPos, logs)
        Exit Function
    End If

    ' Assert non-ASCII subject characters are preserved verbatim inside <title>
    Dim titlePos As Long
    titlePos = InStr(1, html, "<title>", vbTextCompare)
    If titlePos > 0 Then
        Dim titleEndPos As Long
        titleEndPos = InStr(titlePos, html, "</title>", vbTextCompare)
        If titleEndPos > titlePos Then
            Dim titleContent As String
            titleContent = Mid$(html, titlePos + Len("<title>"), titleEndPos - (titlePos + Len("<title>")))
            If InStr(1, titleContent, "á", vbBinaryCompare) = 0 Then
                Test_Issue64Cluster_CuerpoHTMLConEstiloCorporativo_DeclaraCharsetUTF8 = JsonFail("Non-ASCII 'á' lost in <title>. Got: " & titleContent, logs)
                Exit Function
            End If
        End If
    End If

    Test_Issue64Cluster_CuerpoHTMLConEstiloCorporativo_DeclaraCharsetUTF8 = JsonOk("charset_utf8_declared_and_nonascii_preserved", logs)
    Exit Function

EH:
    Test_Issue64Cluster_CuerpoHTMLConEstiloCorporativo_DeclaraCharsetUTF8 = JsonFail("unexpected: " & err.description, logs)
End Function

' ============================================================
' Atom 3 — RechazoPropuestaParaPublicacion_RechazaSinPermisosCalidad (REQ-EML-068-01)
' GIVEN EsCalidad = No (user lacks calidad role)
' WHEN RechazoPropuestaParaPublicacion runs
' THEN err.Number = 1000 raised (or p_Error set)
'   AND TbProyectosEdiciones is NOT modified (no rechazo persisted)
'   AND TbCorreosEnviados has 0 rows
'   AND TbProyectoEdicionesCorreoRevision has 0 rows
' ============================================================
Public Function Test_Issue64Cluster_RechazoPropuestaParaPublicacion_RechazaSinPermisosCalidad() As String
    On Error GoTo EH
    Dim logs(0 To 7) As String
    logs(0) = "1. Arrange: SaveCurrentRoleState; EsCalidad=No, EsAdministrador=No, EsTecnico=No"
    logs(1) = "2. Arrange: SeedAll_Cluster (FIX_EDICION_SI_AVISOS_ID with ParaInformeAvisos=Si)"
    logs(2) = "3. Arrange: LoadTestEdicion + motivo non-empty"
    logs(3) = "4. Act: ed.RechazoPropuestaParaPublicacion"
    logs(4) = "5. Assert: error raised (err.Number=1000 or p_Error set)"
    logs(5) = "6. Assert: count(TbCorreosEnviados) == 0"
    logs(6) = "7. Assert: PropuestaRechazadaPorCalidadMotivo NOT set in TbProyectosEdiciones"
    logs(7) = "8. Teardown: RestoreCurrentRoleState + TeardownAll_Cluster"

    Dim cfgError As String
    If Not Test_Helper.EnsureTestConfigLoaded(cfgError) Then
        Test_Issue64Cluster_RechazoPropuestaParaPublicacion_RechazaSinPermisosCalidad = JsonFail(cfgError, logs)
        Exit Function
    End If

    ' Save role state and force EsCalidad=No
    Dim prevAdmin As EnumSiNo
    Dim prevCalidad As EnumSiNo
    Dim prevTecnico As EnumSiNo
    SaveCurrentRoleState prevAdmin, prevCalidad, prevTecnico
    EsCalidad = EnumSiNo.No
    EsAdministrador = EnumSiNo.No
    EsTecnico = EnumSiNo.No

    SeedAll_Cluster

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Issue64Cluster_RechazoPropuestaParaPublicacion_RechazaSinPermisosCalidad = JsonFail("GetTestDb returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    ' Load edicion
    Dim ed As Edicion
    Dim loadErr As String
    Set ed = Constructor.getEdicion(CStr(FIX_EDICION_SI_AVISOS_ID), loadErr)
    If ed Is Nothing Then
        Test_Issue64Cluster_RechazoPropuestaParaPublicacion_RechazaSinPermisosCalidad = JsonFail("LoadTestEdicion failed: " & loadErr, logs)
        GoTo Teardown
    End If
    ed.PropuestaRechazadaPorCalidadMotivo = "Test rechazo no-calidad"

    ' Act: RechazoPropuestaParaPublicacion
    Dim pError As String
    Dim correoRechazo As correo
    On Error Resume Next
    Set correoRechazo = ed.RechazoPropuestaParaPublicacion(pError)
    Dim errNum As Long
    errNum = err.Number
    On Error GoTo 0

    ' Read DB state
    Dim countCorreos As Long
    Dim countRevision As Long
    Dim motivoPersistido As String
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbCorreosEnviados WHERE IDEdicion=" & FIX_EDICION_SI_AVISOS_ID)
    If Not rs.EOF Then countCorreos = CLng(Nz(rs.fields("C").value, 0))
    rs.Close
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbProyectoEdicionesCorreoRevision WHERE IDEdicion=" & FIX_EDICION_SI_AVISOS_ID)
    If Not rs.EOF Then countRevision = CLng(Nz(rs.fields("C").value, 0))
    rs.Close
    Set rs = db.OpenRecordset("SELECT PropuestaRechazadaPorCalidadMotivo FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_SI_AVISOS_ID)
    If Not rs.EOF Then motivoPersistido = CStr(Nz(rs.fields("PropuestaRechazadaPorCalidadMotivo").value, ""))
    rs.Close
    Set rs = Nothing

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownAll_Cluster
    RestoreCurrentRoleState prevAdmin, prevCalidad, prevTecnico
    Test_Helper.ResetTestSession
    On Error GoTo 0

    ' Assertions
    Dim errorBlocked As Boolean
    errorBlocked = (errNum = 1000) Or (pError <> "")

    If Not errorBlocked Then
        Test_Issue64Cluster_RechazoPropuestaParaPublicacion_RechazaSinPermisosCalidad = JsonFail("Expected error 1000 or p_Error set for non-calidad user. errNum=" & errNum & " pError='" & pError & "'", logs)
        Exit Function
    End If
    If countCorreos <> 0 Then
        Test_Issue64Cluster_RechazoPropuestaParaPublicacion_RechazaSinPermisosCalidad = JsonFail("Expected 0 rows in TbCorreosEnviados but got " & countCorreos, logs)
        Exit Function
    End If
    If countRevision <> 0 Then
        Test_Issue64Cluster_RechazoPropuestaParaPublicacion_RechazaSinPermisosCalidad = JsonFail("Expected 0 rows in TbProyectoEdicionesCorreoRevision but got " & countRevision, logs)
        Exit Function
    End If
    If motivoPersistido <> "" Then
        Test_Issue64Cluster_RechazoPropuestaParaPublicacion_RechazaSinPermisosCalidad = JsonFail("Expected PropuestaRechazadaPorCalidadMotivo empty (rejection NOT persisted) but got '" & motivoPersistido & "'", logs)
        Exit Function
    End If

    Test_Issue64Cluster_RechazoPropuestaParaPublicacion_RechazaSinPermisosCalidad = JsonOk("rechazo_sin_permisos_calidad_bloqueado", logs)
    Exit Function

EH:
    Test_Issue64Cluster_RechazoPropuestaParaPublicacion_RechazaSinPermisosCalidad = JsonFail("unexpected: " & err.description, logs)
    Resume Teardown
End Function
