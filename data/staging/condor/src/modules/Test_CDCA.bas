Attribute VB_Name = "Test_CDCA"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Test_CDCA.bas
' BATERÍA: CDCA - Tests para DatosCDCAServicio.cls
' SKILL:   access-vba-tdd v2.4.2 (migrated from v1.9 legacy)
' BRANCH:  staging
' PROJECT: condor (Dysflow projectId)
'
' MIGRACIÓN v1.9 -> v2.4.2 (CHANGELOG)
'   §1.1.2  Per-test On Error GoTo EH / Catch -> local EH: label + Resume
'           CleanExit. (TestHelper.HandleError será provisto por otra tarea;
'           cross-module On Error GoTo no es posible en VBA, así que seguimos
'           el patrón del exemplar Test_PCSUB_Strict.bas con EH: local.)
'   §3.4    Per-test TestHelper.ForceLocalBackend -> ELIMINADO. SuiteSetup
'           ya invoca BeginTestSession vía TestHelper.SuiteSetup, que setea
'           m_TestingMode + sandbox path + EVE(True). Los tests confían en
'           ese setup a nivel de suite (RunAll llama SuiteSetup al inicio).
'   §4.5    countBefore / countAfter agregados a TODOS los tests que llaman
'           svc.Guardar* (DCount directo sobre TbDatosCDCA) y a PurgaTecnica /
'           Eliminar* (cardinalidad no debe cambiar; sólo mutación de campos).
'   §5.4    db = Nothing -> db = TestHelper.GetTestDb() explícito. Los
'           servicios toman Optional db; pasar Nothing forzaba el camino
'           getdb() interno que NO es lo que queremos testear.
' ==========================================================================

Private Const TEST_ID_BASE As Long = 900000
Private m_TestCounter As Long

Public Sub SuiteSetup(ByRef p_Error As String)
    Call TestHelper.SuiteSetup(p_Error)
End Sub

Public Sub SuiteTeardown(ByRef p_Error As String)
    p_Error = ""
    Call TestHelper.CloseTestDb
    Call TestHelper.RestoreBackend
End Sub

Private Sub SeedAll()
    Dim db As DAO.Database
    Dim idExp As Long
    Dim idSol As Long

    Set db = TestHelper.GetTestDb()
    idExp = GetNextTestId()
    idSol = GetNextTestId()

    db.Execute "INSERT INTO tbExpedientes (idExpediente, CodExp, titulo, objetoContrato, ContratistaPrincipal) " & _
               "VALUES (" & idExp & ", " & TestHelper.SqlStr("EXP-" & idExp) & ", " & TestHelper.SqlStr("TEST EXP") & ", " & TestHelper.SqlStr("TEST OBJ") & ", " & TestHelper.SqlStr("Si") & ")", dbFailOnError

    db.Execute "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, " & _
               "idEstadoInterno, fechaCreacion, usuarioCreacion, fechaModificacion, usuarioModificacion) " & _
               "VALUES (" & idSol & ", " & idExp & ", " & TestHelper.SqlStr("CD/CA") & ", " & TestHelper.SqlStr("CDCA-" & idSol) & ", " & _
               estadoPreregistro & ", Now(), " & TestHelper.SqlStr("TestCDCA") & ", Now(), " & TestHelper.SqlStr("TestCDCA") & ")", dbFailOnError

    db.Execute "INSERT INTO TbDatosCDCA (idDatosCDCA, idSolicitud) VALUES (" & GetNextTestId() & ", " & idSol & ")", dbFailOnError
End Sub

Private Function Seed_X() As Long
    Dim db As DAO.Database
    Dim idExp As Long
    Dim idSol As Long

    Set db = TestHelper.GetTestDb()
    idExp = GetNextTestId()
    idSol = GetNextTestId()

    db.Execute "INSERT INTO tbExpedientes (idExpediente, CodExp, titulo, objetoContrato, ContratistaPrincipal) " & _
               "VALUES (" & idExp & ", " & TestHelper.SqlStr("EXP-" & idExp) & ", " & TestHelper.SqlStr("TEST EXP") & ", " & TestHelper.SqlStr("TEST OBJ") & ", " & TestHelper.SqlStr("Si") & ")", dbFailOnError

    db.Execute "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, " & _
               "idEstadoInterno, fechaCreacion, usuarioCreacion, fechaModificacion, usuarioModificacion) " & _
               "VALUES (" & idSol & ", " & idExp & ", " & TestHelper.SqlStr("CD/CA") & ", " & TestHelper.SqlStr("CDCA-" & idSol) & ", " & _
               estadoPreregistro & ", Now(), " & TestHelper.SqlStr("TestCDCA") & ", Now(), " & TestHelper.SqlStr("TestCDCA") & ")", dbFailOnError

    db.Execute "INSERT INTO TbDatosCDCA (idDatosCDCA, idSolicitud) VALUES (" & GetNextTestId() & ", " & idSol & ")", dbFailOnError

    Seed_X = idSol
End Function

Private Function SeedCompleteCDCAFixture() As Long
    Dim db As DAO.Database
    Dim idExp As Long
    Dim idSol As Long

    Set db = TestHelper.GetTestDb()
    idExp = GetNextTestId()
    idSol = GetNextTestId()

    db.Execute "INSERT INTO tbExpedientes (idExpediente, CodExp, titulo, objetoContrato, ContratistaPrincipal) " & _
               "VALUES (" & idExp & ", " & TestHelper.SqlStr("EXP-" & idExp) & ", " & TestHelper.SqlStr("TEST EXP") & ", " & TestHelper.SqlStr("TEST OBJ") & ", " & TestHelper.SqlStr("Si") & ")", dbFailOnError

    db.Execute "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, " & _
               "idEstadoInterno, fechaCreacion, usuarioCreacion, fechaModificacion, usuarioModificacion) " & _
               "VALUES (" & idSol & ", " & idExp & ", " & TestHelper.SqlStr("CD/CA") & ", " & TestHelper.SqlStr("CDCA-" & idSol) & ", " & _
               estadoPreregistro & ", Now(), " & TestHelper.SqlStr("TestCDCA") & ", Now(), " & TestHelper.SqlStr("TestCDCA") & ")", dbFailOnError

    db.Execute "INSERT INTO TbDatosCDCA (idDatosCDCA, idSolicitud, " & _
               "numContrato, refSuministrador, SuministradorNombreDir, refDesviacionesPrevias, requiereModificacionContrato, " & _
               "identificacionMaterial, numPlanoEspecificacion, cantidadPeriodo, numSerieLote, causaNC, " & _
               "descripcionImpactoNC, descripcionImpactoNCCont, " & _
               "afectaPrestaciones, afectaSeguridad, afectaFiabilidad, afectaVidaUtil, afectaMedioambiente, " & _
               "afectaMantenibilidad, afectaIntercambiabilidad, afectaApariencia, afectaOtros, " & _
               "impactoCoste, clasificacionNC, esSuministradorAD, identificacionAutoridadDiseno, efectoFechaEntrega, " & _
               "firmaAprobacionRespIngenieriaNombre, firmaAprobacionRespProduccionNombre, firmaAprobacionRespCalidadNombre, " & _
               "firmaAprobacionRespDisenioNombre, firmaAprobacionRepresentanteSumNombre, " & _
               "racCodigo, observacionesRAC, racNombre, racDecision, " & _
               "decisionFinal, observacionesFinales, NombreFirmanteFinal) " & _
               "VALUES (" & GetNextTestId() & ", " & idSol & ", " & _
               TestHelper.SqlStr("CONTRATO-TEST") & ", " & TestHelper.SqlStr("REF-SUM-TEST") & ", " & TestHelper.SqlStr("SUMINISTRADOR TEST") & ", " & _
               TestHelper.SqlStr("REF-DEV-001") & ", True, " & _
               TestHelper.SqlStr("MATERIAL-001") & ", " & TestHelper.SqlStr("PLANO-001") & ", " & TestHelper.SqlStr("PER-001") & ", " & TestHelper.SqlStr("SERIE-001") & ", " & TestHelper.SqlStr("CAUSA NC") & ", " & _
               TestHelper.SqlStr("IMPACTO DESCRIPCION") & ", " & TestHelper.SqlStr("IMPACTO CONTINUACION") & ", " & _
               "True, True, True, True, True, True, True, True, True, " & _
               TestHelper.SqlStr("AUMENTARA") & ", " & TestHelper.SqlStr("MAYOR") & ", True, " & TestHelper.SqlStr("AUT-DIS-001") & ", " & TestHelper.SqlStr("EFECTO FECHA") & ", " & _
               TestHelper.SqlStr("FIRMA INGENIERIA") & ", " & TestHelper.SqlStr("FIRMA PRODUCCION") & ", " & TestHelper.SqlStr("FIRMA CALIDAD") & ", " & _
               TestHelper.SqlStr("FIRMA DISENIO") & ", " & TestHelper.SqlStr("FIRMA REPRESENTANTE") & ", " & _
               TestHelper.SqlStr("RAC-001") & ", " & TestHelper.SqlStr("OBS RAC") & ", " & TestHelper.SqlStr("NOMBRE RAC") & ", " & TestHelper.SqlStr("APROBADO") & ", " & _
               TestHelper.SqlStr("DECISION FINAL") & ", " & TestHelper.SqlStr("OBSERVACIONES FINALES") & ", " & TestHelper.SqlStr("NOMBRE FIRMANTE") & ")", dbFailOnError

    SeedCompleteCDCAFixture = idSol
End Function

Private Function SeedTecnicaFixture() As Long
    Dim db As DAO.Database
    Dim idExp As Long
    Dim idSol As Long

    Set db = TestHelper.GetTestDb()
    idExp = GetNextTestId()
    idSol = GetNextTestId()

    db.Execute "INSERT INTO tbExpedientes (idExpediente, CodExp, titulo, objetoContrato, ContratistaPrincipal) " & _
               "VALUES (" & idExp & ", " & TestHelper.SqlStr("EXP-" & idExp) & ", " & TestHelper.SqlStr("TEST EXP") & ", " & TestHelper.SqlStr("TEST OBJ") & ", " & TestHelper.SqlStr("Si") & ")", dbFailOnError

    db.Execute "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, " & _
               "idEstadoInterno, fechaCreacion, usuarioCreacion, fechaModificacion, usuarioModificacion) " & _
               "VALUES (" & idSol & ", " & idExp & ", " & TestHelper.SqlStr("CD/CA") & ", " & TestHelper.SqlStr("CDCA-" & idSol) & ", " & _
               estadoPreregistro & ", Now(), " & TestHelper.SqlStr("TestCDCA") & ", Now(), " & TestHelper.SqlStr("TestCDCA") & ")", dbFailOnError

    db.Execute "INSERT INTO TbDatosCDCA (idDatosCDCA, idSolicitud, " & _
               "identificacionMaterial, numPlanoEspecificacion, cantidadPeriodo, numSerieLote, causaNC, " & _
               "descripcionImpactoNC, descripcionImpactoNCCont, " & _
               "afectaPrestaciones, afectaSeguridad, afectaFiabilidad, afectaVidaUtil, afectaMedioambiente, " & _
               "afectaMantenibilidad, afectaIntercambiabilidad, afectaApariencia, afectaOtros, " & _
               "impactoCoste, clasificacionNC, esSuministradorAD, identificacionAutoridadDiseno, efectoFechaEntrega) " & _
               "VALUES (" & GetNextTestId() & ", " & idSol & ", " & _
               TestHelper.SqlStr("MATERIAL-001") & ", " & TestHelper.SqlStr("PLANO-001") & ", " & TestHelper.SqlStr("PER-001") & ", " & TestHelper.SqlStr("SERIE-001") & ", " & TestHelper.SqlStr("CAUSA NC") & ", " & _
               TestHelper.SqlStr("IMPACTO DESCRIPCION") & ", " & TestHelper.SqlStr("IMPACTO CONTINUACION") & ", " & _
               "True, True, True, True, True, True, True, True, True, " & _
               TestHelper.SqlStr("AUMENTARA") & ", " & TestHelper.SqlStr("MAYOR") & ", True, " & TestHelper.SqlStr("AUT-DIS-001") & ", " & TestHelper.SqlStr("EFECTO FECHA") & ")", dbFailOnError

    SeedTecnicaFixture = idSol
End Function

Public Sub TeardownAll()
    Dim db As DAO.Database
    Set db = TestHelper.GetTestDb()
    On Error Resume Next
    db.Execute "DELETE FROM TbDatosCDCA WHERE idSolicitud >= " & TEST_ID_BASE, dbFailOnError
    db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud >= " & TEST_ID_BASE, dbFailOnError
    db.Execute "DELETE FROM tbExpedientes WHERE idExpediente >= " & TEST_ID_BASE, dbFailOnError
End Sub

Private Function GetNextTestId() As Long
    m_TestCounter = m_TestCounter + 1
    GetNextTestId = TEST_ID_BASE + m_TestCounter
End Function

Public Function RunAll() As String
    Dim runError As String
    Dim logs(0 To 5) As String

    Call SuiteSetup(runError)
    If runError <> "" Then
        Dim failLogs(0 To 0) As String
        failLogs(0) = "SuiteSetup fallo: " & runError
        RunAll = TestHelper.BuildJsonFail("TESTS BLOCKED", failLogs)
        Exit Function
    End If
    logs(0) = "SuiteSetup OK"

    Call SeedAll
    m_TestCounter = 0

    Dim passed As Long: passed = 0
    Dim failed As Long: failed = 0
    Dim resultJson As String

    ' GROUP 1
    resultJson = Test_GetDatosCDCA_Happy():        If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_GetDatosCDCA_NotFound():    If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_GetDatosCDCA_NegativeId():  If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_GetDatosCDCA_UnknownId():  If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1

    ' GROUP 2
    resultJson = Test_EsDatosGeneralesCompleta_True():                         If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_EsDatosGeneralesCompleta_MissingNumContrato():            If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_EsDatosGeneralesCompleta_MissingRefSuministrador():      If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_EsDatosGeneralesCompleta_MissingSuministradorNombreDir(): If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_EsDatosGeneralesCompleta_MissingRequiereModificacion():  If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1

    ' GROUP 3
    resultJson = Test_EsParteTecnicaCompleta_True():              If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_EsParteTecnicaCompleta_MissingMaterial():  If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_EsParteTecnicaCompleta_MissingCausaNC():  If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_EsParteTecnicaCompleta_MissingDescripcion(): If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_EsParteTecnicaCompleta_MissingCoste():    If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_EsParteTecnicaCompleta_MissingClasificacion(): If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_EsParteTecnicaCompleta_MissingEsSuministradorAD(): If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1

    ' GROUP 4
    resultJson = Test_EsDetalleCompleto_True():       If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_EsDetalleCompleto_MissingMaterial(): If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_EsDetalleCompleto_MissingCausaNC(): If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1

    ' GROUP 5
    resultJson = Test_EsMotivosCompleto_True():      If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_EsMotivosCompleto_MissingDescripcion(): If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1
    resultJson = Test_EsMotivosCompleto_MissingDescripcionCont(): If Left(resultJson, 7) = "{""ok"":true" Then passed = passed + 1 Else failed = failed + 1

    Call TeardownAll

    Dim teardownError As String
    Call SuiteTeardown(teardownError)
    logs(1) = "SuiteTeardown OK"

    RunAll = "{""ok"":true,""passed"":" & passed & ",""failed"":" & failed & ",""logs"":[" & _
             TestHelper.EscapeJsonString(logs(0)) & "," & TestHelper.EscapeJsonString(logs(1)) & "]}"
End Function

' ==========================================================================
' GROUP 1 - GetDatosCDCA (4 tests)
' ==========================================================================

Public Function Test_GetDatosCDCA_Happy() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio
    Dim result As DatosCDCA

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    logs(0) = "1. Seeded idSol=" & idSol

    Set result = svc.GetDatosCDCA(idSol, db)
    If Not result Is Nothing Then
        logs(1) = "2. GetDatosCDCA returned non-Nothing result"
        Test_GetDatosCDCA_Happy = TestHelper.BuildJsonOk("DatosCDCA", logs)
    Else
        logs(1) = "2. FAILED: result is Nothing for valid idSol"
        Test_GetDatosCDCA_Happy = TestHelper.BuildJsonFail("Expected DatosCDCA, got Nothing", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GetDatosCDCA_Happy: ERR " & Err.Number & " - " & Err.Description
    Test_GetDatosCDCA_Happy = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GetDatosCDCA_NotFound() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim svc As New DatosCDCAServicio
    Dim result As DatosCDCA
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    logs(0) = "1. id=0 raises Err 513 or returns Nothing (both acceptable)"

    On Error Resume Next
    Set result = svc.GetDatosCDCA(0, db)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber = 513 Then
        logs(1) = "2. GetDatosCDCA raised Err 513 for id=0"
        Test_GetDatosCDCA_NotFound = TestHelper.BuildJsonOk("Nothing", logs)
    ElseIf result Is Nothing Then
        logs(1) = "2. GetDatosCDCA correctly returned Nothing for id=0"
        Test_GetDatosCDCA_NotFound = TestHelper.BuildJsonOk("Nothing", logs)
    Else
        logs(1) = "2. FAILED: expected Nothing or Err 513 for id=0"
        Test_GetDatosCDCA_NotFound = TestHelper.BuildJsonFail("Expected Nothing for id=0", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GetDatosCDCA_NotFound: ERR " & Err.Number & " - " & Err.Description
    Test_GetDatosCDCA_NotFound = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GetDatosCDCA_NegativeId() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim svc As New DatosCDCAServicio
    Dim result As DatosCDCA
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    logs(0) = "1. id=-1 raises Err 513 (idSolicitud <= 0)"

    On Error Resume Next
    Set result = svc.GetDatosCDCA(-1, db)
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber = 513 Then
        logs(1) = "2. GetDatosCDCA raised Err 513 for id=-1"
        Test_GetDatosCDCA_NegativeId = TestHelper.BuildJsonOk("Nothing", logs)
    Else
        logs(1) = "2. FAILED: expected Err 513 for id=-1, got Err " & errNumber
        Test_GetDatosCDCA_NegativeId = TestHelper.BuildJsonFail("Expected Err 513 for id=-1", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GetDatosCDCA_NegativeId: ERR " & Err.Number & " - " & Err.Description
    Test_GetDatosCDCA_NegativeId = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GetDatosCDCA_UnknownId() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim svc As New DatosCDCAServicio
    Dim result As DatosCDCA

    Set db = TestHelper.GetTestDb()
    logs(0) = "1. id=999999 returns Nothing (no error)"

    Set result = svc.GetDatosCDCA(999999, db)
    If result Is Nothing Then
        logs(1) = "2. GetDatosCDCA correctly returned Nothing for unknown id"
        Test_GetDatosCDCA_UnknownId = TestHelper.BuildJsonOk("Nothing", logs)
    Else
        logs(1) = "2. FAILED: expected Nothing for id=999999"
        Test_GetDatosCDCA_UnknownId = TestHelper.BuildJsonFail("Expected Nothing for id=999999", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GetDatosCDCA_UnknownId: ERR " & Err.Number & " - " & Err.Description
    Test_GetDatosCDCA_UnknownId = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 2 - EsDatosGeneralesCompleta (5 tests)
' ==========================================================================

Public Function Test_EsDatosGeneralesCompleta_True() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedCompleteCDCAFixture()
    logs(0) = "1. Seeded complete fixture idSol=" & idSol

    If svc.EsDatosGeneralesCompleta(idSol, db) Then
        logs(1) = "2. EsDatosGeneralesCompleta=True as expected"
        Test_EsDatosGeneralesCompleta_True = TestHelper.BuildJsonOk("true", logs)
    Else
        logs(1) = "2. FAILED: EsDatosGeneralesCompleta=False on complete fixture"
        Test_EsDatosGeneralesCompleta_True = TestHelper.BuildJsonFail("EsDatosGeneralesCompleta returned False for complete fixture", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsDatosGeneralesCompleta_True: ERR " & Err.Number & " - " & Err.Description
    Test_EsDatosGeneralesCompleta_True = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsDatosGeneralesCompleta_MissingNumContrato() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    db.Execute "UPDATE TbDatosCDCA SET numContrato=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with numContrato=NULL"

    If Not svc.EsDatosGeneralesCompleta(idSol, db) Then
        logs(1) = "2. EsDatosGeneralesCompleta=False as expected"
        Test_EsDatosGeneralesCompleta_MissingNumContrato = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsDatosGeneralesCompleta=True for missing numContrato"
        Test_EsDatosGeneralesCompleta_MissingNumContrato = TestHelper.BuildJsonFail("EsDatosGeneralesCompleta returned True for missing numContrato", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsDatosGeneralesCompleta_MissingNumContrato: ERR " & Err.Number & " - " & Err.Description
    Test_EsDatosGeneralesCompleta_MissingNumContrato = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsDatosGeneralesCompleta_MissingRefSuministrador() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    db.Execute "UPDATE TbDatosCDCA SET refSuministrador=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with refSuministrador=NULL"

    If Not svc.EsDatosGeneralesCompleta(idSol, db) Then
        logs(1) = "2. EsDatosGeneralesCompleta=False as expected"
        Test_EsDatosGeneralesCompleta_MissingRefSuministrador = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsDatosGeneralesCompleta=True for missing refSuministrador"
        Test_EsDatosGeneralesCompleta_MissingRefSuministrador = TestHelper.BuildJsonFail("EsDatosGeneralesCompleta returned True for missing refSuministrador", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsDatosGeneralesCompleta_MissingRefSuministrador: ERR " & Err.Number & " - " & Err.Description
    Test_EsDatosGeneralesCompleta_MissingRefSuministrador = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsDatosGeneralesCompleta_MissingSuministradorNombreDir() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    db.Execute "UPDATE TbDatosCDCA SET SuministradorNombreDir=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with SuministradorNombreDir=NULL"

    If Not svc.EsDatosGeneralesCompleta(idSol, db) Then
        logs(1) = "2. EsDatosGeneralesCompleta=False as expected"
        Test_EsDatosGeneralesCompleta_MissingSuministradorNombreDir = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsDatosGeneralesCompleta=True for missing SuministradorNombreDir"
        Test_EsDatosGeneralesCompleta_MissingSuministradorNombreDir = TestHelper.BuildJsonFail("EsDatosGeneralesCompleta returned True for missing SuministradorNombreDir", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsDatosGeneralesCompleta_MissingSuministradorNombreDir: ERR " & Err.Number & " - " & Err.Description
    Test_EsDatosGeneralesCompleta_MissingSuministradorNombreDir = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsDatosGeneralesCompleta_MissingRequiereModificacion() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    db.Execute "UPDATE TbDatosCDCA SET requiereModificacionContrato=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with requiereModificacionContrato=NULL"

    If Not svc.EsDatosGeneralesCompleta(idSol, db) Then
        logs(1) = "2. EsDatosGeneralesCompleta=False as expected"
        Test_EsDatosGeneralesCompleta_MissingRequiereModificacion = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsDatosGeneralesCompleta=True for missing requiereModificacionContrato"
        Test_EsDatosGeneralesCompleta_MissingRequiereModificacion = TestHelper.BuildJsonFail("EsDatosGeneralesCompleta returned True for missing requiereModificacionContrato", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsDatosGeneralesCompleta_MissingRequiereModificacion: ERR " & Err.Number & " - " & Err.Description
    Test_EsDatosGeneralesCompleta_MissingRequiereModificacion = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 3 - EsParteTecnicaCompleta (7 tests)
' ==========================================================================

Public Function Test_EsParteTecnicaCompleta_True() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedTecnicaFixture()
    logs(0) = "1. Seeded tecnica fixture idSol=" & idSol

    If svc.EsParteTecnicaCompleta(idSol, db) Then
        logs(1) = "2. EsParteTecnicaCompleta=True as expected"
        Test_EsParteTecnicaCompleta_True = TestHelper.BuildJsonOk("true", logs)
    Else
        logs(1) = "2. FAILED: EsParteTecnicaCompleta=False on complete tecnica"
        Test_EsParteTecnicaCompleta_True = TestHelper.BuildJsonFail("EsParteTecnicaCompleta returned False for complete tecnica fixture", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsParteTecnicaCompleta_True: ERR " & Err.Number & " - " & Err.Description
    Test_EsParteTecnicaCompleta_True = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsParteTecnicaCompleta_MissingMaterial() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedTecnicaFixture()
    db.Execute "UPDATE TbDatosCDCA SET identificacionMaterial=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with identificacionMaterial=NULL"

    If Not svc.EsParteTecnicaCompleta(idSol, db) Then
        logs(1) = "2. EsParteTecnicaCompleta=False as expected"
        Test_EsParteTecnicaCompleta_MissingMaterial = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsParteTecnicaCompleta=True for missing material"
        Test_EsParteTecnicaCompleta_MissingMaterial = TestHelper.BuildJsonFail("EsParteTecnicaCompleta returned True for missing material", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsParteTecnicaCompleta_MissingMaterial: ERR " & Err.Number & " - " & Err.Description
    Test_EsParteTecnicaCompleta_MissingMaterial = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsParteTecnicaCompleta_MissingCausaNC() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedTecnicaFixture()
    db.Execute "UPDATE TbDatosCDCA SET causaNC=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with causaNC=NULL"

    If Not svc.EsParteTecnicaCompleta(idSol, db) Then
        logs(1) = "2. EsParteTecnicaCompleta=False as expected"
        Test_EsParteTecnicaCompleta_MissingCausaNC = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsParteTecnicaCompleta=True for missing causaNC"
        Test_EsParteTecnicaCompleta_MissingCausaNC = TestHelper.BuildJsonFail("EsParteTecnicaCompleta returned True for missing causaNC", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsParteTecnicaCompleta_MissingCausaNC: ERR " & Err.Number & " - " & Err.Description
    Test_EsParteTecnicaCompleta_MissingCausaNC = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsParteTecnicaCompleta_MissingDescripcion() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedTecnicaFixture()
    db.Execute "UPDATE TbDatosCDCA SET descripcionImpactoNC=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with descripcionImpactoNC=NULL"

    If Not svc.EsParteTecnicaCompleta(idSol, db) Then
        logs(1) = "2. EsParteTecnicaCompleta=False as expected"
        Test_EsParteTecnicaCompleta_MissingDescripcion = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsParteTecnicaCompleta=True for missing descripcionImpactoNC"
        Test_EsParteTecnicaCompleta_MissingDescripcion = TestHelper.BuildJsonFail("EsParteTecnicaCompleta returned True for missing descripcionImpactoNC", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsParteTecnicaCompleta_MissingDescripcion: ERR " & Err.Number & " - " & Err.Description
    Test_EsParteTecnicaCompleta_MissingDescripcion = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsParteTecnicaCompleta_MissingCoste() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedTecnicaFixture()
    db.Execute "UPDATE TbDatosCDCA SET impactoCoste=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with impactoCoste=NULL"

    If Not svc.EsParteTecnicaCompleta(idSol, db) Then
        logs(1) = "2. EsParteTecnicaCompleta=False as expected"
        Test_EsParteTecnicaCompleta_MissingCoste = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsParteTecnicaCompleta=True for missing impactoCoste"
        Test_EsParteTecnicaCompleta_MissingCoste = TestHelper.BuildJsonFail("EsParteTecnicaCompleta returned True for missing impactoCoste", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsParteTecnicaCompleta_MissingCoste: ERR " & Err.Number & " - " & Err.Description
    Test_EsParteTecnicaCompleta_MissingCoste = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsParteTecnicaCompleta_MissingClasificacion() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedTecnicaFixture()
    db.Execute "UPDATE TbDatosCDCA SET clasificacionNC=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with clasificacionNC=NULL"

    If Not svc.EsParteTecnicaCompleta(idSol, db) Then
        logs(1) = "2. EsParteTecnicaCompleta=False as expected"
        Test_EsParteTecnicaCompleta_MissingClasificacion = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsParteTecnicaCompleta=True for missing clasificacionNC"
        Test_EsParteTecnicaCompleta_MissingClasificacion = TestHelper.BuildJsonFail("EsParteTecnicaCompleta returned True for missing clasificacionNC", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsParteTecnicaCompleta_MissingClasificacion: ERR " & Err.Number & " - " & Err.Description
    Test_EsParteTecnicaCompleta_MissingClasificacion = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsParteTecnicaCompleta_MissingEsSuministradorAD() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedTecnicaFixture()
    db.Execute "UPDATE TbDatosCDCA SET esSuministradorAD=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with esSuministradorAD=NULL"

    If Not svc.EsParteTecnicaCompleta(idSol, db) Then
        logs(1) = "2. EsParteTecnicaCompleta=False as expected"
        Test_EsParteTecnicaCompleta_MissingEsSuministradorAD = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsParteTecnicaCompleta=True for null esSuministradorAD"
        Test_EsParteTecnicaCompleta_MissingEsSuministradorAD = TestHelper.BuildJsonFail("EsParteTecnicaCompleta returned True for null esSuministradorAD", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsParteTecnicaCompleta_MissingEsSuministradorAD: ERR " & Err.Number & " - " & Err.Description
    Test_EsParteTecnicaCompleta_MissingEsSuministradorAD = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 4 - EsDetalleCompleto (3 tests)
' ==========================================================================

Public Function Test_EsDetalleCompleto_True() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedTecnicaFixture()
    logs(0) = "1. Seeded tecnica fixture idSol=" & idSol

    If svc.EsDetalleCompleto(idSol, db) Then
        logs(1) = "2. EsDetalleCompleto=True as expected"
        Test_EsDetalleCompleto_True = TestHelper.BuildJsonOk("true", logs)
    Else
        logs(1) = "2. FAILED: EsDetalleCompleto=False on complete fixture"
        Test_EsDetalleCompleto_True = TestHelper.BuildJsonFail("EsDetalleCompleto returned False for complete fixture", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsDetalleCompleto_True: ERR " & Err.Number & " - " & Err.Description
    Test_EsDetalleCompleto_True = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsDetalleCompleto_MissingMaterial() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedTecnicaFixture()
    db.Execute "UPDATE TbDatosCDCA SET identificacionMaterial=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with identificacionMaterial=NULL"

    If Not svc.EsDetalleCompleto(idSol, db) Then
        logs(1) = "2. EsDetalleCompleto=False as expected"
        Test_EsDetalleCompleto_MissingMaterial = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsDetalleCompleto=True for missing material"
        Test_EsDetalleCompleto_MissingMaterial = TestHelper.BuildJsonFail("EsDetalleCompleto returned True for missing material", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsDetalleCompleto_MissingMaterial: ERR " & Err.Number & " - " & Err.Description
    Test_EsDetalleCompleto_MissingMaterial = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsDetalleCompleto_MissingCausaNC() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedTecnicaFixture()
    db.Execute "UPDATE TbDatosCDCA SET causaNC=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with causaNC=NULL"

    If Not svc.EsDetalleCompleto(idSol, db) Then
        logs(1) = "2. EsDetalleCompleto=False as expected"
        Test_EsDetalleCompleto_MissingCausaNC = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsDetalleCompleto=True for missing causaNC"
        Test_EsDetalleCompleto_MissingCausaNC = TestHelper.BuildJsonFail("EsDetalleCompleto returned True for missing causaNC", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsDetalleCompleto_MissingCausaNC: ERR " & Err.Number & " - " & Err.Description
    Test_EsDetalleCompleto_MissingCausaNC = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 5 - EsMotivosCompleto (3 tests)
' ==========================================================================

Public Function Test_EsMotivosCompleto_True() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedTecnicaFixture()
    logs(0) = "1. Seeded tecnica fixture idSol=" & idSol

    If svc.EsMotivosCompleto(idSol, db) Then
        logs(1) = "2. EsMotivosCompleto=True as expected"
        Test_EsMotivosCompleto_True = TestHelper.BuildJsonOk("true", logs)
    Else
        logs(1) = "2. FAILED: EsMotivosCompleto=False on complete motivos"
        Test_EsMotivosCompleto_True = TestHelper.BuildJsonFail("EsMotivosCompleto returned False for complete fixture", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsMotivosCompleto_True: ERR " & Err.Number & " - " & Err.Description
    Test_EsMotivosCompleto_True = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsMotivosCompleto_MissingDescripcion() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedTecnicaFixture()
    db.Execute "UPDATE TbDatosCDCA SET descripcionImpactoNC=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with descripcionImpactoNC=NULL"

    If Not svc.EsMotivosCompleto(idSol, db) Then
        logs(1) = "2. EsMotivosCompleto=False as expected"
        Test_EsMotivosCompleto_MissingDescripcion = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsMotivosCompleto=True for missing descripcionImpactoNC"
        Test_EsMotivosCompleto_MissingDescripcion = TestHelper.BuildJsonFail("EsMotivosCompleto returned True for missing descripcionImpactoNC", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsMotivosCompleto_MissingDescripcion: ERR " & Err.Number & " - " & Err.Description
    Test_EsMotivosCompleto_MissingDescripcion = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsMotivosCompleto_MissingDescripcionCont() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedTecnicaFixture()
    db.Execute "UPDATE TbDatosCDCA SET descripcionImpactoNCCont=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with descripcionImpactoNCCont=NULL"

    If Not svc.EsMotivosCompleto(idSol, db) Then
        logs(1) = "2. EsMotivosCompleto=False as expected"
        Test_EsMotivosCompleto_MissingDescripcionCont = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsMotivosCompleto=True for missing descripcionImpactoNCCont"
        Test_EsMotivosCompleto_MissingDescripcionCont = TestHelper.BuildJsonFail("EsMotivosCompleto returned True for missing descripcionImpactoNCCont", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsMotivosCompleto_MissingDescripcionCont: ERR " & Err.Number & " - " & Err.Description
    Test_EsMotivosCompleto_MissingDescripcionCont = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 6 - EsDictamenRACCompleta (3 tests)
' ==========================================================================

Public Function Test_EsDictamenRACCompleta_True() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedCompleteCDCAFixture()
    logs(0) = "1. Seeded complete fixture idSol=" & idSol

    If svc.EsDictamenRACCompleto(idSol, db) Then
        logs(1) = "2. EsDictamenRACCompleto=True as expected"
        Test_EsDictamenRACCompleta_True = TestHelper.BuildJsonOk("true", logs)
    Else
        logs(1) = "2. FAILED: EsDictamenRACCompleto=False on complete RAC"
        Test_EsDictamenRACCompleta_True = TestHelper.BuildJsonFail("EsDictamenRACCompleto returned False for complete fixture", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsDictamenRACCompleta_True: ERR " & Err.Number & " - " & Err.Description
    Test_EsDictamenRACCompleta_True = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsDictamenRACCompleta_MissingRACCodigo() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedCompleteCDCAFixture()
    db.Execute "UPDATE TbDatosCDCA SET racCodigo=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with racCodigo=NULL"

    If Not svc.EsDictamenRACCompleto(idSol, db) Then
        logs(1) = "2. EsDictamenRACCompleto=False as expected"
        Test_EsDictamenRACCompleta_MissingRACCodigo = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsDictamenRACCompleto=True for missing racCodigo"
        Test_EsDictamenRACCompleta_MissingRACCodigo = TestHelper.BuildJsonFail("EsDictamenRACCompleto returned True for missing racCodigo", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsDictamenRACCompleta_MissingRACCodigo: ERR " & Err.Number & " - " & Err.Description
    Test_EsDictamenRACCompleta_MissingRACCodigo = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsDictamenRACCompleta_MissingRACDecision() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedCompleteCDCAFixture()
    db.Execute "UPDATE TbDatosCDCA SET racDecision=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with racDecision=NULL"

    If Not svc.EsDictamenRACCompleto(idSol, db) Then
        logs(1) = "2. EsDictamenRACCompleto=False as expected"
        Test_EsDictamenRACCompleta_MissingRACDecision = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsDictamenRACCompleto=True for missing racDecision"
        Test_EsDictamenRACCompleta_MissingRACDecision = TestHelper.BuildJsonFail("EsDictamenRACCompleto returned True for missing racDecision", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsDictamenRACCompleta_MissingRACDecision: ERR " & Err.Number & " - " & Err.Description
    Test_EsDictamenRACCompleta_MissingRACDecision = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 7 - EsAprobacionSuministradorCompleta (3 tests)
' ==========================================================================

Public Function Test_EsAprobacionSuministradorCompleta_True() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedCompleteCDCAFixture()
    logs(0) = "1. Seeded complete fixture idSol=" & idSol

    If svc.EsAprobacionSuministradorCompleta(idSol, db) Then
        logs(1) = "2. EsAprobacionSuministradorCompleta=True as expected"
        Test_EsAprobacionSuministradorCompleta_True = TestHelper.BuildJsonOk("true", logs)
    Else
        logs(1) = "2. FAILED: EsAprobacionSuministradorCompleta=False on complete aprobacion"
        Test_EsAprobacionSuministradorCompleta_True = TestHelper.BuildJsonFail("EsAprobacionSuministradorCompleta returned False for complete fixture", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsAprobacionSuministradorCompleta_True: ERR " & Err.Number & " - " & Err.Description
    Test_EsAprobacionSuministradorCompleta_True = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsAprobacionSuministradorCompleta_MissingFirmaIngenieria() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedCompleteCDCAFixture()
    db.Execute "UPDATE TbDatosCDCA SET firmaAprobacionRespIngenieriaNombre=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with firmaAprobacionRespIngenieriaNombre=NULL"

    If Not svc.EsAprobacionSuministradorCompleta(idSol, db) Then
        logs(1) = "2. EsAprobacionSuministradorCompleta=False as expected"
        Test_EsAprobacionSuministradorCompleta_MissingFirmaIngenieria = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsAprobacionSuministradorCompleta=True for missing firmaAprobacionRespIngenieriaNombre"
        Test_EsAprobacionSuministradorCompleta_MissingFirmaIngenieria = TestHelper.BuildJsonFail("EsAprobacionSuministradorCompleta returned True for missing firmaAprobacionRespIngenieriaNombre", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsAprobacionSuministradorCompleta_MissingFirmaIngenieria: ERR " & Err.Number & " - " & Err.Description
    Test_EsAprobacionSuministradorCompleta_MissingFirmaIngenieria = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsAprobacionSuministradorCompleta_MissingFirmaProduccion() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedCompleteCDCAFixture()
    db.Execute "UPDATE TbDatosCDCA SET firmaAprobacionRespProduccionNombre=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with firmaAprobacionRespProduccionNombre=NULL"

    If Not svc.EsAprobacionSuministradorCompleta(idSol, db) Then
        logs(1) = "2. EsAprobacionSuministradorCompleta=False as expected"
        Test_EsAprobacionSuministradorCompleta_MissingFirmaProduccion = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsAprobacionSuministradorCompleta=True for missing firmaAprobacionRespProduccionNombre"
        Test_EsAprobacionSuministradorCompleta_MissingFirmaProduccion = TestHelper.BuildJsonFail("EsAprobacionSuministradorCompleta returned True for missing firmaAprobacionRespProduccionNombre", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsAprobacionSuministradorCompleta_MissingFirmaProduccion: ERR " & Err.Number & " - " & Err.Description
    Test_EsAprobacionSuministradorCompleta_MissingFirmaProduccion = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 8 - EsDecisionFinalCompleta (3 tests)
' ==========================================================================

Public Function Test_EsDecisionFinalCompleta_True() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedCompleteCDCAFixture()
    logs(0) = "1. Seeded complete fixture idSol=" & idSol

    If svc.EsDecisionFinalCompleta(idSol, db) Then
        logs(1) = "2. EsDecisionFinalCompleta=True as expected"
        Test_EsDecisionFinalCompleta_True = TestHelper.BuildJsonOk("true", logs)
    Else
        logs(1) = "2. FAILED: EsDecisionFinalCompleta=False on complete decision"
        Test_EsDecisionFinalCompleta_True = TestHelper.BuildJsonFail("EsDecisionFinalCompleta returned False for complete fixture", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsDecisionFinalCompleta_True: ERR " & Err.Number & " - " & Err.Description
    Test_EsDecisionFinalCompleta_True = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsDecisionFinalCompleta_MissingDecision() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedCompleteCDCAFixture()
    db.Execute "UPDATE TbDatosCDCA SET decisionFinal=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with decisionFinal=NULL"

    If Not svc.EsDecisionFinalCompleta(idSol, db) Then
        logs(1) = "2. EsDecisionFinalCompleta=False as expected"
        Test_EsDecisionFinalCompleta_MissingDecision = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsDecisionFinalCompleta=True for missing decisionFinal"
        Test_EsDecisionFinalCompleta_MissingDecision = TestHelper.BuildJsonFail("EsDecisionFinalCompleta returned True for missing decisionFinal", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsDecisionFinalCompleta_MissingDecision: ERR " & Err.Number & " - " & Err.Description
    Test_EsDecisionFinalCompleta_MissingDecision = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_EsDecisionFinalCompleta_MissingFirmante() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim svc As New DatosCDCAServicio

    Set db = TestHelper.GetTestDb()
    idSol = SeedCompleteCDCAFixture()
    db.Execute "UPDATE TbDatosCDCA SET NombreFirmanteFinal=NULL WHERE idSolicitud=" & idSol, dbFailOnError
    logs(0) = "1. Seeded idSol=" & idSol & " with NombreFirmanteFinal=NULL"

    If Not svc.EsDecisionFinalCompleta(idSol, db) Then
        logs(1) = "2. EsDecisionFinalCompleta=False as expected"
        Test_EsDecisionFinalCompleta_MissingFirmante = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: EsDecisionFinalCompleta=True for missing NombreFirmanteFinal"
        Test_EsDecisionFinalCompleta_MissingFirmante = TestHelper.BuildJsonFail("EsDecisionFinalCompleta returned True for missing NombreFirmanteFinal", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EsDecisionFinalCompleta_MissingFirmante: ERR " & Err.Number & " - " & Err.Description
    Test_EsDecisionFinalCompleta_MissingFirmante = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 9 - PurgaTecnica (2 tests)
' ==========================================================================

Public Function Test_PurgaTecnica_Happy() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim rs As DAO.Recordset
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = SeedTecnicaFixture()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded tecnica idSol=" & idSol & " countBefore=" & countBefore

    On Error Resume Next
    svc.PurgaTecnica idSol, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: PurgaTecnica raised Err " & errNumber
        Test_PurgaTecnica_Happy = TestHelper.BuildJsonFail("PurgaTecnica raised: " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. PurgaTecnica succeeded"

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(2) = "3. FAILED: cardinality changed before=" & countBefore & " after=" & countAfter
        Test_PurgaTecnica_Happy = TestHelper.BuildJsonFail("PurgaTecnica changed row count", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected " & countBefore & ")"

    Set rs = db.OpenRecordset("SELECT identificacionMaterial, causaNC, descripcionImpactoNC FROM TbDatosCDCA WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If rs.EOF Then
        logs(3) = "4. FAILED: row not found"
        Test_PurgaTecnica_Happy = TestHelper.BuildJsonFail("TbDatosCDCA row not found", logs)
        GoTo CleanExit
    End If
    If Not (IsNull(rs!identificacionMaterial) And IsNull(rs!causaNC) And IsNull(rs!descripcionImpactoNC)) Then
        logs(3) = "4. FAILED: tecnica fields not cleared"
        Test_PurgaTecnica_Happy = TestHelper.BuildJsonFail("PurgaTecnica did not clear tecnica fields", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. tecnica fields correctly cleared"

    logs(4) = "5. PASS"
    Test_PurgaTecnica_Happy = TestHelper.BuildJsonOk("true", logs)
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PurgaTecnica_Happy: ERR " & Err.Number & " - " & Err.Description
    Test_PurgaTecnica_Happy = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_PurgaTecnica_NotFound() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim svc As New DatosCDCAServicio
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    logs(0) = "1. id=0 raises Err 513 (idSolicitud <= 0)"

    On Error Resume Next
    svc.PurgaTecnica 0, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber = 513 Then
        logs(1) = "2. PurgaTecnica raised Err 513 for id=0 (expected)"
        Test_PurgaTecnica_NotFound = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: expected Err 513 for id=0, got Err " & errNumber
        Test_PurgaTecnica_NotFound = TestHelper.BuildJsonFail("Expected Err 513 for id=0", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PurgaTecnica_NotFound: ERR " & Err.Number & " - " & Err.Description
    Test_PurgaTecnica_NotFound = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 10 - EliminarDictamenRAC (1 test)
' ==========================================================================

Public Function Test_EliminarDictamenRAC_Happy() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim rs As DAO.Recordset
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = SeedCompleteCDCAFixture()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded complete idSol=" & idSol & " countBefore=" & countBefore

    On Error Resume Next
    svc.EliminarDictamenRAC idSol, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: EliminarDictamenRAC raised Err " & errNumber
        Test_EliminarDictamenRAC_Happy = TestHelper.BuildJsonFail("EliminarDictamenRAC raised: " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. EliminarDictamenRAC succeeded"

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(2) = "3. FAILED: cardinality changed"
        Test_EliminarDictamenRAC_Happy = TestHelper.BuildJsonFail("EliminarDictamenRAC changed row count", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected " & countBefore & ")"

    Set rs = db.OpenRecordset("SELECT racCodigo, racDecision FROM TbDatosCDCA WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If rs.EOF Then
        logs(3) = "4. FAILED: row not found"
        Test_EliminarDictamenRAC_Happy = TestHelper.BuildJsonFail("TbDatosCDCA row not found", logs)
        GoTo CleanExit
    End If
    If Not (IsNull(rs!racCodigo) And IsNull(rs!racDecision)) Then
        logs(3) = "4. FAILED: RAC fields not cleared"
        Test_EliminarDictamenRAC_Happy = TestHelper.BuildJsonFail("EliminarDictamenRAC did not clear RAC fields", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. RAC fields correctly cleared"

    logs(4) = "5. PASS"
    Test_EliminarDictamenRAC_Happy = TestHelper.BuildJsonOk("true", logs)
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EliminarDictamenRAC_Happy: ERR " & Err.Number & " - " & Err.Description
    Test_EliminarDictamenRAC_Happy = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 11 - EliminarAprobacionSuministrador (1 test)
' ==========================================================================

Public Function Test_EliminarAprobacionSuministrador_Happy() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim rs As DAO.Recordset
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = SeedCompleteCDCAFixture()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded complete idSol=" & idSol & " countBefore=" & countBefore

    On Error Resume Next
    svc.EliminarAprobacionSuministrador idSol, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: EliminarAprobacionSuministrador raised Err " & errNumber
        Test_EliminarAprobacionSuministrador_Happy = TestHelper.BuildJsonFail("EliminarAprobacionSuministrador raised: " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. EliminarAprobacionSuministrador succeeded"

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(2) = "3. FAILED: cardinality changed"
        Test_EliminarAprobacionSuministrador_Happy = TestHelper.BuildJsonFail("EliminarAprobacionSuministrador changed row count", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected " & countBefore & ")"

    Set rs = db.OpenRecordset("SELECT firmaAprobacionRespIngenieriaNombre, firmaAprobacionRespProduccionNombre FROM TbDatosCDCA WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If rs.EOF Then
        logs(3) = "4. FAILED: row not found"
        Test_EliminarAprobacionSuministrador_Happy = TestHelper.BuildJsonFail("TbDatosCDCA row not found", logs)
        GoTo CleanExit
    End If
    If Not (IsNull(rs!firmaAprobacionRespIngenieriaNombre) And IsNull(rs!firmaAprobacionRespProduccionNombre)) Then
        logs(3) = "4. FAILED: firma fields not cleared"
        Test_EliminarAprobacionSuministrador_Happy = TestHelper.BuildJsonFail("EliminarAprobacionSuministrador did not clear firma fields", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. firma fields correctly cleared"

    logs(4) = "5. PASS"
    Test_EliminarAprobacionSuministrador_Happy = TestHelper.BuildJsonOk("true", logs)
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EliminarAprobacionSuministrador_Happy: ERR " & Err.Number & " - " & Err.Description
    Test_EliminarAprobacionSuministrador_Happy = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 12 - EliminarDecisionFinal (1 test)
' ==========================================================================

Public Function Test_EliminarDecisionFinal_Happy() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim rs As DAO.Recordset
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = SeedCompleteCDCAFixture()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded complete idSol=" & idSol & " countBefore=" & countBefore

    On Error Resume Next
    svc.EliminarDecisionFinal idSol, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: EliminarDecisionFinal raised Err " & errNumber
        Test_EliminarDecisionFinal_Happy = TestHelper.BuildJsonFail("EliminarDecisionFinal raised: " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. EliminarDecisionFinal succeeded"

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(2) = "3. FAILED: cardinality changed"
        Test_EliminarDecisionFinal_Happy = TestHelper.BuildJsonFail("EliminarDecisionFinal changed row count", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected " & countBefore & ")"

    Set rs = db.OpenRecordset("SELECT decisionFinal, NombreFirmanteFinal FROM TbDatosCDCA WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If rs.EOF Then
        logs(3) = "4. FAILED: row not found"
        Test_EliminarDecisionFinal_Happy = TestHelper.BuildJsonFail("TbDatosCDCA row not found", logs)
        GoTo CleanExit
    End If
    If Not (IsNull(rs!decisionFinal) And IsNull(rs!NombreFirmanteFinal)) Then
        logs(3) = "4. FAILED: decision fields not cleared"
        Test_EliminarDecisionFinal_Happy = TestHelper.BuildJsonFail("EliminarDecisionFinal did not clear decision fields", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. decision fields correctly cleared"

    logs(4) = "5. PASS"
    Test_EliminarDecisionFinal_Happy = TestHelper.BuildJsonOk("true", logs)
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_EliminarDecisionFinal_Happy: ERR " & Err.Number & " - " & Err.Description
    Test_EliminarDecisionFinal_Happy = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 14 - GuardarDatosGenerales (4 tests)
' ==========================================================================

Public Function Test_GuardarDatosGenerales_Happy() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim rs As DAO.Recordset

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.numContrato = "CONTRATO-TEST-001"
    datos.refSuministrador = "REF-SUM-001"
    datos.SuministradorNombreDir = "Suministrador Test SA"
    datos.requiereModificacionContrato = True
    datos.idSolicitud = idSol

    If svc.GuardarDatosGenerales(datos, db) Then
        countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
        If countAfter <> countBefore Then
            logs(1) = "2. FAILED: cardinality changed before=" & countBefore & " after=" & countAfter
            Test_GuardarDatosGenerales_Happy = TestHelper.BuildJsonFail("GuardarDatosGenerales changed row count", logs)
            GoTo CleanExit
        End If
        logs(1) = "2. GuardarDatosGenerales=True; countAfter=" & countAfter

        Set rs = db.OpenRecordset("SELECT numContrato FROM TbDatosCDCA WHERE idSolicitud=" & idSol, dbOpenSnapshot)
        If rs.EOF Or Nz(rs!numContrato, "") <> "CONTRATO-TEST-001" Then
            logs(2) = "3. FAILED: numContrato not persisted"
            Test_GuardarDatosGenerales_Happy = TestHelper.BuildJsonFail("numContrato not found", logs)
            GoTo CleanExit
        End If
        logs(2) = "3. numContrato persisted"

        logs(3) = "4. PASS"
        Test_GuardarDatosGenerales_Happy = TestHelper.BuildJsonOk("true", logs)
    Else
        logs(1) = "2. FAILED: GuardarDatosGenerales returned False"
        Test_GuardarDatosGenerales_Happy = TestHelper.BuildJsonFail("GuardarDatosGenerales returned False", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarDatosGenerales_Happy: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarDatosGenerales_Happy = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarDatosGenerales_NotFound() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.numContrato = "CONTRATO-TEST"
    datos.idSolicitud = idSol

    If Not svc.GuardarDatosGenerales(datos, db) Then
        countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
        If countAfter <> countBefore Then
            logs(1) = "2. FAILED: cardinality changed"
            Test_GuardarDatosGenerales_NotFound = TestHelper.BuildJsonFail("row count changed on False", logs)
            GoTo CleanExit
        End If
        logs(1) = "2. GuardarDatosGenerales=False; countAfter=" & countAfter
        Test_GuardarDatosGenerales_NotFound = TestHelper.BuildJsonOk("false", logs)
    Else
        logs(1) = "2. FAILED: GuardarDatosGenerales=True (expected False)"
        Test_GuardarDatosGenerales_NotFound = TestHelper.BuildJsonFail("GuardarDatosGenerales returned True", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarDatosGenerales_NotFound: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarDatosGenerales_NotFound = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarDatosGenerales_Update() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim rs As DAO.Recordset

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.numContrato = "CONTRATO-V1"
    datos.refSuministrador = "REF-V1"
    datos.SuministradorNombreDir = "Suministrador V1"
    datos.requiereModificacionContrato = False
    datos.idSolicitud = idSol

    If svc.GuardarDatosGenerales(datos, db) Then
        datos.numContrato = "CONTRATO-V2"
        If svc.GuardarDatosGenerales(datos, db) Then
            countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
            If countAfter <> countBefore Then
                logs(1) = "2. FAILED: cardinality changed"
                Test_GuardarDatosGenerales_Update = TestHelper.BuildJsonFail("Update changed row count", logs)
                GoTo CleanExit
            End If
            logs(1) = "2. Two calls succeeded; countAfter=" & countAfter

            Set rs = db.OpenRecordset("SELECT numContrato FROM TbDatosCDCA WHERE idSolicitud=" & idSol, dbOpenSnapshot)
            If rs.EOF Or Nz(rs!numContrato, "") <> "CONTRATO-V2" Then
                logs(2) = "3. FAILED: not updated to V2"
                Test_GuardarDatosGenerales_Update = TestHelper.BuildJsonFail("Second GuardarDatosGenerales did not update", logs)
                GoTo CleanExit
            End If
            logs(2) = "3. updated to V2"

            logs(3) = "4. PASS"
            Test_GuardarDatosGenerales_Update = TestHelper.BuildJsonOk("true", logs)
        Else
            logs(1) = "2. FAILED: second call returned False"
            Test_GuardarDatosGenerales_Update = TestHelper.BuildJsonFail("Second GuardarDatosGenerales returned False", logs)
        End If
    Else
        logs(1) = "2. FAILED: first call returned False"
        Test_GuardarDatosGenerales_Update = TestHelper.BuildJsonFail("First GuardarDatosGenerales returned False", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarDatosGenerales_Update: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarDatosGenerales_Update = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarDatosGenerales_WithoutRequiredFields() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.numContrato = ""
    datos.refSuministrador = "REF-SUM"
    datos.SuministradorNombreDir = "Suministrador"
    datos.requiereModificacionContrato = True
    datos.idSolicitud = idSol

    On Error Resume Next
    svc.GuardarDatosGenerales datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(1) = "2. FAILED: cardinality changed"
        Test_GuardarDatosGenerales_WithoutRequiredFields = TestHelper.BuildJsonFail("cardinality changed on validation error", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. countAfter=" & countAfter & " (no side effect; validation Err " & errNumber & ")"
    Test_GuardarDatosGenerales_WithoutRequiredFields = TestHelper.BuildJsonOk("false", logs)
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarDatosGenerales_WithoutRequiredFields: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarDatosGenerales_WithoutRequiredFields = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 15 - GuardarPropuesta (4 tests)
' ==========================================================================

Public Function Test_GuardarPropuesta_Happy() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim rs As DAO.Recordset
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.identificacionMaterial = "MATERIAL-TEST"
    datos.numPlanoEspecificacion = "PLANO-TEST"
    datos.cantidadPeriodo = "PER-TEST"
    datos.numSerieLote = "SERIE-TEST"
    datos.idSolicitud = idSol

    On Error Resume Next
    svc.GuardarPropuesta datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: GuardarPropuesta raised Err " & errNumber
        Test_GuardarPropuesta_Happy = TestHelper.BuildJsonFail("GuardarPropuesta raised: " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarPropuesta succeeded"

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(2) = "3. FAILED: cardinality changed"
        Test_GuardarPropuesta_Happy = TestHelper.BuildJsonFail("GuardarPropuesta changed row count", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected " & countBefore & ")"

    Set rs = db.OpenRecordset("SELECT identificacionMaterial FROM TbDatosCDCA WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If rs.EOF Or Nz(rs!identificacionMaterial, "") <> "MATERIAL-TEST" Then
        logs(3) = "4. FAILED: identificacionMaterial not persisted"
        Test_GuardarPropuesta_Happy = TestHelper.BuildJsonFail("field not found", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. identificacionMaterial persisted"

    logs(4) = "5. PASS"
    Test_GuardarPropuesta_Happy = TestHelper.BuildJsonOk("true", logs)
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarPropuesta_Happy: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarPropuesta_Happy = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarPropuesta_NotFound() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim countBefore As Long
    Dim countAfter As Long
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=0")
    logs(0) = "1. countBefore=" & countBefore & " (id=0 baseline)"

    datos.identificacionMaterial = "MATERIAL-TEST"
    datos.numPlanoEspecificacion = "PLANO-TEST"
    datos.cantidadPeriodo = "PER-TEST"
    datos.numSerieLote = "SERIE-TEST"
    datos.idSolicitud = 0

    On Error Resume Next
    svc.GuardarPropuesta datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=0")
    If countAfter <> countBefore Then
        logs(1) = "2. FAILED: cardinality changed before=" & countBefore & " after=" & countAfter
        Test_GuardarPropuesta_NotFound = TestHelper.BuildJsonFail("cardinality changed", logs)
        GoTo CleanExit
    End If
    If errNumber = 0 Then
        logs(1) = "2. FAILED: expected error for id=0, got none"
        Test_GuardarPropuesta_NotFound = TestHelper.BuildJsonFail("Expected error for id=0", logs)
    Else
        logs(1) = "2. GuardarPropuesta raised Err " & errNumber & " (expected); countAfter=" & countAfter
        Test_GuardarPropuesta_NotFound = TestHelper.BuildJsonOk("false", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarPropuesta_NotFound: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarPropuesta_NotFound = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarPropuesta_Update() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim rs As DAO.Recordset
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.identificacionMaterial = "MAT-V1"
    datos.numPlanoEspecificacion = "PLANO-V1"
    datos.cantidadPeriodo = "PER-V1"
    datos.numSerieLote = "SERIE-V1"
    datos.idSolicitud = idSol

    On Error Resume Next
    svc.GuardarPropuesta datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: first call raised Err " & errNumber
        Test_GuardarPropuesta_Update = TestHelper.BuildJsonFail("First GuardarPropuesta raised: " & errNumber, logs)
        GoTo CleanExit
    End If

    datos.identificacionMaterial = "MAT-V2"
    On Error Resume Next
    svc.GuardarPropuesta datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: second call raised Err " & errNumber
        Test_GuardarPropuesta_Update = TestHelper.BuildJsonFail("Second GuardarPropuesta raised: " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. Both calls succeeded"

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(2) = "3. FAILED: cardinality changed"
        Test_GuardarPropuesta_Update = TestHelper.BuildJsonFail("Update changed row count", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected " & countBefore & ")"

    Set rs = db.OpenRecordset("SELECT identificacionMaterial FROM TbDatosCDCA WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If rs.EOF Or Nz(rs!identificacionMaterial, "") <> "MAT-V2" Then
        logs(3) = "4. FAILED: not updated to V2"
        Test_GuardarPropuesta_Update = TestHelper.BuildJsonFail("Second GuardarPropuesta did not update", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. updated to V2"

    logs(4) = "5. PASS"
    Test_GuardarPropuesta_Update = TestHelper.BuildJsonOk("true", logs)
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarPropuesta_Update: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarPropuesta_Update = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarPropuesta_WithoutRequiredFields() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.identificacionMaterial = ""
    datos.numPlanoEspecificacion = "PLANO-TEST"
    datos.cantidadPeriodo = "PER-TEST"
    datos.numSerieLote = "SERIE-TEST"
    datos.idSolicitud = idSol

    On Error Resume Next
    svc.GuardarPropuesta datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(1) = "2. FAILED: cardinality changed"
        Test_GuardarPropuesta_WithoutRequiredFields = TestHelper.BuildJsonFail("cardinality changed on validation error", logs)
        GoTo CleanExit
    End If
    If errNumber = 0 Then
        logs(1) = "2. FAILED: expected validation error"
        Test_GuardarPropuesta_WithoutRequiredFields = TestHelper.BuildJsonFail("Expected validation error", logs)
    Else
        logs(1) = "2. GuardarPropuesta raised Err " & errNumber & " (validation); countAfter=" & countAfter
        Test_GuardarPropuesta_WithoutRequiredFields = TestHelper.BuildJsonOk("false", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarPropuesta_WithoutRequiredFields: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarPropuesta_WithoutRequiredFields = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 16 - GuardarImpacto (4 tests)
' ==========================================================================

Public Function Test_GuardarImpacto_Happy() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim rs As DAO.Recordset
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.causaNC = "CAUSA-NC-TEST"
    datos.descripcionImpactoNC = "DESC IMPACTO"
    datos.descripcionImpactoNCCont = "CONTINUACION"
    datos.impactoCoste = "AUMENTARA"
    datos.clasificacionNC = "MAYOR"
    datos.esSuministradorAD = True
    datos.identificacionAutoridadDiseno = "AUT-DIS-001"
    datos.efectoFechaEntrega = "EFECTO FECHA"
    datos.idSolicitud = idSol

    On Error Resume Next
    svc.GuardarImpacto datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: GuardarImpacto raised Err " & errNumber
        Test_GuardarImpacto_Happy = TestHelper.BuildJsonFail("GuardarImpacto raised: " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarImpacto succeeded"

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(2) = "3. FAILED: cardinality changed"
        Test_GuardarImpacto_Happy = TestHelper.BuildJsonFail("GuardarImpacto changed row count", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected " & countBefore & ")"

    Set rs = db.OpenRecordset("SELECT causaNC FROM TbDatosCDCA WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If rs.EOF Or Nz(rs!causaNC, "") <> "CAUSA-NC-TEST" Then
        logs(3) = "4. FAILED: causaNC not persisted"
        Test_GuardarImpacto_Happy = TestHelper.BuildJsonFail("field not found", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. causaNC persisted"

    logs(4) = "5. PASS"
    Test_GuardarImpacto_Happy = TestHelper.BuildJsonOk("true", logs)
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarImpacto_Happy: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarImpacto_Happy = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarImpacto_NotFound() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim countBefore As Long
    Dim countAfter As Long
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=0")
    logs(0) = "1. countBefore=" & countBefore & " (id=0 baseline)"

    datos.causaNC = "CAUSA-NC-TEST"
    datos.descripcionImpactoNC = "DESC"
    datos.descripcionImpactoNCCont = "CONT"
    datos.impactoCoste = "AUMENTARA"
    datos.clasificacionNC = "MAYOR"
    datos.esSuministradorAD = True
    datos.identificacionAutoridadDiseno = "AUT-1"
    datos.efectoFechaEntrega = "EFECTO"
    datos.idSolicitud = 0

    On Error Resume Next
    svc.GuardarImpacto datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=0")
    If countAfter <> countBefore Then
        logs(1) = "2. FAILED: cardinality changed"
        Test_GuardarImpacto_NotFound = TestHelper.BuildJsonFail("cardinality changed", logs)
        GoTo CleanExit
    End If
    If errNumber = 0 Then
        logs(1) = "2. FAILED: expected error for id=0"
        Test_GuardarImpacto_NotFound = TestHelper.BuildJsonFail("Expected error for id=0", logs)
    Else
        logs(1) = "2. GuardarImpacto raised Err " & errNumber & " (expected); countAfter=" & countAfter
        Test_GuardarImpacto_NotFound = TestHelper.BuildJsonOk("false", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarImpacto_NotFound: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarImpacto_NotFound = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarImpacto_Update() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim rs As DAO.Recordset
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.causaNC = "CAUSA-V1"
    datos.descripcionImpactoNC = "DESC V1"
    datos.descripcionImpactoNCCont = "CONT V1"
    datos.impactoCoste = "AUMENTARA"
    datos.clasificacionNC = "MAYOR"
    datos.esSuministradorAD = True
    datos.identificacionAutoridadDiseno = "AUT-1"
    datos.efectoFechaEntrega = "EFECTO V1"
    datos.idSolicitud = idSol

    On Error Resume Next
    svc.GuardarImpacto datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: first call raised Err " & errNumber
        Test_GuardarImpacto_Update = TestHelper.BuildJsonFail("First GuardarImpacto raised: " & errNumber, logs)
        GoTo CleanExit
    End If

    datos.causaNC = "CAUSA-V2"
    On Error Resume Next
    svc.GuardarImpacto datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: second call raised Err " & errNumber
        Test_GuardarImpacto_Update = TestHelper.BuildJsonFail("Second GuardarImpacto raised: " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. Both calls succeeded"

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(2) = "3. FAILED: cardinality changed"
        Test_GuardarImpacto_Update = TestHelper.BuildJsonFail("Update changed row count", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected " & countBefore & ")"

    Set rs = db.OpenRecordset("SELECT causaNC FROM TbDatosCDCA WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If rs.EOF Or Nz(rs!causaNC, "") <> "CAUSA-V2" Then
        logs(3) = "4. FAILED: not updated to V2"
        Test_GuardarImpacto_Update = TestHelper.BuildJsonFail("Second GuardarImpacto did not update", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. updated to V2"

    logs(4) = "5. PASS"
    Test_GuardarImpacto_Update = TestHelper.BuildJsonOk("true", logs)
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarImpacto_Update: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarImpacto_Update = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarImpacto_WithoutRequiredFields() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.causaNC = ""
    datos.descripcionImpactoNC = "DESC"
    datos.descripcionImpactoNCCont = "CONT"
    datos.impactoCoste = "AUMENTARA"
    datos.clasificacionNC = "MAYOR"
    datos.esSuministradorAD = True
    datos.identificacionAutoridadDiseno = "AUT-1"
    datos.efectoFechaEntrega = "EFECTO"
    datos.idSolicitud = idSol

    On Error Resume Next
    svc.GuardarImpacto datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(1) = "2. FAILED: cardinality changed"
        Test_GuardarImpacto_WithoutRequiredFields = TestHelper.BuildJsonFail("cardinality changed on validation error", logs)
        GoTo CleanExit
    End If
    If errNumber = 0 Then
        logs(1) = "2. FAILED: expected validation error"
        Test_GuardarImpacto_WithoutRequiredFields = TestHelper.BuildJsonFail("Expected validation error", logs)
    Else
        logs(1) = "2. GuardarImpacto raised Err " & errNumber & " (validation); countAfter=" & countAfter
        Test_GuardarImpacto_WithoutRequiredFields = TestHelper.BuildJsonOk("false", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarImpacto_WithoutRequiredFields: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarImpacto_WithoutRequiredFields = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 17 - GuardarAprobacionSuministrador (4 tests)
' ==========================================================================

Public Function Test_GuardarAprobacionSuministrador_Happy() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim rs As DAO.Recordset
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.firmaAprobacionRespIngenieriaNombre = "FIRMA ING"
    datos.firmaAprobacionRespProduccionNombre = "FIRMA PROD"
    datos.firmaAprobacionRespCalidadNombre = "FIRMA CAL"
    datos.firmaAprobacionRespDisenioNombre = "FIRMA DIS"
    datos.firmaAprobacionRepresentanteSumNombre = "FIRMA REP"
    datos.idSolicitud = idSol

    On Error Resume Next
    svc.GuardarAprobacionSuministrador datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: GuardarAprobacionSuministrador raised Err " & errNumber
        Test_GuardarAprobacionSuministrador_Happy = TestHelper.BuildJsonFail("GuardarAprobacionSuministrador raised: " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarAprobacionSuministrador succeeded"

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(2) = "3. FAILED: cardinality changed"
        Test_GuardarAprobacionSuministrador_Happy = TestHelper.BuildJsonFail("GuardarAprobacionSuministrador changed row count", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected " & countBefore & ")"

    Set rs = db.OpenRecordset("SELECT firmaAprobacionRespIngenieriaNombre FROM TbDatosCDCA WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If rs.EOF Or Nz(rs!firmaAprobacionRespIngenieriaNombre, "") <> "FIRMA ING" Then
        logs(3) = "4. FAILED: firma fields not persisted"
        Test_GuardarAprobacionSuministrador_Happy = TestHelper.BuildJsonFail("field not found", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. firmaAprobacionRespIngenieriaNombre persisted"

    logs(4) = "5. PASS"
    Test_GuardarAprobacionSuministrador_Happy = TestHelper.BuildJsonOk("true", logs)
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarAprobacionSuministrador_Happy: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarAprobacionSuministrador_Happy = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarAprobacionSuministrador_NotFound() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim countBefore As Long
    Dim countAfter As Long
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=0")
    logs(0) = "1. countBefore=" & countBefore & " (id=0 baseline)"

    datos.firmaAprobacionRespIngenieriaNombre = "FIRMA ING"
    datos.firmaAprobacionRespCalidadNombre = "FIRMA CAL"
    datos.idSolicitud = 0

    On Error Resume Next
    svc.GuardarAprobacionSuministrador datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=0")
    If countAfter <> countBefore Then
        logs(1) = "2. FAILED: cardinality changed"
        Test_GuardarAprobacionSuministrador_NotFound = TestHelper.BuildJsonFail("cardinality changed", logs)
        GoTo CleanExit
    End If
    If errNumber = 0 Then
        logs(1) = "2. FAILED: expected error for id=0"
        Test_GuardarAprobacionSuministrador_NotFound = TestHelper.BuildJsonFail("Expected error for id=0", logs)
    Else
        logs(1) = "2. GuardarAprobacionSuministrador raised Err " & errNumber & " (expected); countAfter=" & countAfter
        Test_GuardarAprobacionSuministrador_NotFound = TestHelper.BuildJsonOk("false", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarAprobacionSuministrador_NotFound: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarAprobacionSuministrador_NotFound = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarAprobacionSuministrador_Update() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim rs As DAO.Recordset
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.firmaAprobacionRespIngenieriaNombre = "FIRMA V1"
    datos.firmaAprobacionRespProduccionNombre = "FIRMA PROD V1"
    datos.firmaAprobacionRespCalidadNombre = "FIRMA CAL V1"
    datos.firmaAprobacionRespDisenioNombre = "FIRMA DIS V1"
    datos.firmaAprobacionRepresentanteSumNombre = "FIRMA REP V1"
    datos.idSolicitud = idSol

    On Error Resume Next
    svc.GuardarAprobacionSuministrador datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: first call raised Err " & errNumber
        Test_GuardarAprobacionSuministrador_Update = TestHelper.BuildJsonFail("First GuardarAprobacionSuministrador raised: " & errNumber, logs)
        GoTo CleanExit
    End If

    datos.firmaAprobacionRespIngenieriaNombre = "FIRMA V2"
    On Error Resume Next
    svc.GuardarAprobacionSuministrador datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: second call raised Err " & errNumber
        Test_GuardarAprobacionSuministrador_Update = TestHelper.BuildJsonFail("Second GuardarAprobacionSuministrador raised: " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. Both calls succeeded"

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(2) = "3. FAILED: cardinality changed"
        Test_GuardarAprobacionSuministrador_Update = TestHelper.BuildJsonFail("Update changed row count", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected " & countBefore & ")"

    Set rs = db.OpenRecordset("SELECT firmaAprobacionRespIngenieriaNombre FROM TbDatosCDCA WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If rs.EOF Or Nz(rs!firmaAprobacionRespIngenieriaNombre, "") <> "FIRMA V2" Then
        logs(3) = "4. FAILED: not updated to V2"
        Test_GuardarAprobacionSuministrador_Update = TestHelper.BuildJsonFail("Second GuardarAprobacionSuministrador did not update", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. updated to V2"

    logs(4) = "5. PASS"
    Test_GuardarAprobacionSuministrador_Update = TestHelper.BuildJsonOk("true", logs)
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarAprobacionSuministrador_Update: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarAprobacionSuministrador_Update = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarAprobacionSuministrador_WithoutRequiredFields() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.firmaAprobacionRespIngenieriaNombre = ""
    datos.firmaAprobacionRespCalidadNombre = "FIRMA CAL"
    datos.idSolicitud = idSol

    On Error Resume Next
    svc.GuardarAprobacionSuministrador datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(1) = "2. FAILED: cardinality changed"
        Test_GuardarAprobacionSuministrador_WithoutRequiredFields = TestHelper.BuildJsonFail("cardinality changed on validation error", logs)
        GoTo CleanExit
    End If
    If errNumber = 0 Then
        logs(1) = "2. FAILED: expected validation error"
        Test_GuardarAprobacionSuministrador_WithoutRequiredFields = TestHelper.BuildJsonFail("Expected validation error", logs)
    Else
        logs(1) = "2. GuardarAprobacionSuministrador raised Err " & errNumber & " (validation); countAfter=" & countAfter
        Test_GuardarAprobacionSuministrador_WithoutRequiredFields = TestHelper.BuildJsonOk("false", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarAprobacionSuministrador_WithoutRequiredFields: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarAprobacionSuministrador_WithoutRequiredFields = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 18 - GuardarDictamenRAC (4 tests)
' ==========================================================================

Public Function Test_GuardarDictamenRAC_Happy() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim rs As DAO.Recordset
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.racCodigo = "RAC-TEST-001"
    datos.observacionesRAC = "OBS RAC TEST"
    datos.racNombre = "NOMBRE RAC"
    datos.racDecision = "APROBADO"
    datos.idSolicitud = idSol

    On Error Resume Next
    svc.GuardarDictamenRAC datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: GuardarDictamenRAC raised Err " & errNumber
        Test_GuardarDictamenRAC_Happy = TestHelper.BuildJsonFail("GuardarDictamenRAC raised: " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarDictamenRAC succeeded"

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(2) = "3. FAILED: cardinality changed"
        Test_GuardarDictamenRAC_Happy = TestHelper.BuildJsonFail("GuardarDictamenRAC changed row count", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected " & countBefore & ")"

    Set rs = db.OpenRecordset("SELECT racCodigo FROM TbDatosCDCA WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If rs.EOF Or Nz(rs!racCodigo, "") <> "RAC-TEST-001" Then
        logs(3) = "4. FAILED: racCodigo not persisted"
        Test_GuardarDictamenRAC_Happy = TestHelper.BuildJsonFail("field not found", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. racCodigo persisted"

    logs(4) = "5. PASS"
    Test_GuardarDictamenRAC_Happy = TestHelper.BuildJsonOk("true", logs)
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarDictamenRAC_Happy: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarDictamenRAC_Happy = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarDictamenRAC_NotFound() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim countBefore As Long
    Dim countAfter As Long
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=0")
    logs(0) = "1. countBefore=" & countBefore & " (id=0 baseline)"

    datos.racCodigo = "RAC-TEST"
    datos.idSolicitud = 0

    On Error Resume Next
    svc.GuardarDictamenRAC datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=0")
    If countAfter <> countBefore Then
        logs(1) = "2. FAILED: cardinality changed"
        Test_GuardarDictamenRAC_NotFound = TestHelper.BuildJsonFail("cardinality changed", logs)
        GoTo CleanExit
    End If
    If errNumber = 0 Then
        logs(1) = "2. FAILED: expected error for id=0"
        Test_GuardarDictamenRAC_NotFound = TestHelper.BuildJsonFail("Expected error for id=0", logs)
    Else
        logs(1) = "2. GuardarDictamenRAC raised Err " & errNumber & " (expected); countAfter=" & countAfter
        Test_GuardarDictamenRAC_NotFound = TestHelper.BuildJsonOk("false", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarDictamenRAC_NotFound: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarDictamenRAC_NotFound = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarDictamenRAC_Update() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim rs As DAO.Recordset
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.racCodigo = "RAC-V1"
    datos.observacionesRAC = "OBS V1"
    datos.racNombre = "NOMBRE V1"
    datos.racDecision = "APROBADO"
    datos.idSolicitud = idSol

    On Error Resume Next
    svc.GuardarDictamenRAC datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: first call raised Err " & errNumber
        Test_GuardarDictamenRAC_Update = TestHelper.BuildJsonFail("First GuardarDictamenRAC raised: " & errNumber, logs)
        GoTo CleanExit
    End If

    datos.racCodigo = "RAC-V2"
    On Error Resume Next
    svc.GuardarDictamenRAC datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: second call raised Err " & errNumber
        Test_GuardarDictamenRAC_Update = TestHelper.BuildJsonFail("Second GuardarDictamenRAC raised: " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. Both calls succeeded"

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(2) = "3. FAILED: cardinality changed"
        Test_GuardarDictamenRAC_Update = TestHelper.BuildJsonFail("Update changed row count", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected " & countBefore & ")"

    Set rs = db.OpenRecordset("SELECT racCodigo FROM TbDatosCDCA WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If rs.EOF Or Nz(rs!racCodigo, "") <> "RAC-V2" Then
        logs(3) = "4. FAILED: not updated to V2"
        Test_GuardarDictamenRAC_Update = TestHelper.BuildJsonFail("Second GuardarDictamenRAC did not update", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. updated to V2"

    logs(4) = "5. PASS"
    Test_GuardarDictamenRAC_Update = TestHelper.BuildJsonOk("true", logs)
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarDictamenRAC_Update: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarDictamenRAC_Update = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarDictamenRAC_WithoutRequiredFields() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.racCodigo = ""
    datos.observacionesRAC = "OBS"
    datos.racNombre = "NOMBRE"
    datos.racDecision = "APROBADO"
    datos.idSolicitud = idSol

    On Error Resume Next
    svc.GuardarDictamenRAC datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(1) = "2. FAILED: cardinality changed"
        Test_GuardarDictamenRAC_WithoutRequiredFields = TestHelper.BuildJsonFail("cardinality changed on validation error", logs)
        GoTo CleanExit
    End If
    If errNumber = 0 Then
        logs(1) = "2. FAILED: expected validation error"
        Test_GuardarDictamenRAC_WithoutRequiredFields = TestHelper.BuildJsonFail("Expected validation error", logs)
    Else
        logs(1) = "2. GuardarDictamenRAC raised Err " & errNumber & " (validation); countAfter=" & countAfter
        Test_GuardarDictamenRAC_WithoutRequiredFields = TestHelper.BuildJsonOk("false", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarDictamenRAC_WithoutRequiredFields: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarDictamenRAC_WithoutRequiredFields = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' GROUP 19 - GuardarDecisionFinal (4 tests)
' ==========================================================================

Public Function Test_GuardarDecisionFinal_Happy() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim rs As DAO.Recordset
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.decisionFinal = "DECISION FINAL TEST"
    datos.observacionesFinales = "OBS FINALES"
    datos.NombreFirmanteFinal = "FIRMANTE FINAL"
    datos.idSolicitud = idSol

    On Error Resume Next
    svc.GuardarDecisionFinal datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: GuardarDecisionFinal raised Err " & errNumber
        Test_GuardarDecisionFinal_Happy = TestHelper.BuildJsonFail("GuardarDecisionFinal raised: " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. GuardarDecisionFinal succeeded"

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(2) = "3. FAILED: cardinality changed"
        Test_GuardarDecisionFinal_Happy = TestHelper.BuildJsonFail("GuardarDecisionFinal changed row count", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected " & countBefore & ")"

    Set rs = db.OpenRecordset("SELECT decisionFinal FROM TbDatosCDCA WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If rs.EOF Or Nz(rs!decisionFinal, "") <> "DECISION FINAL TEST" Then
        logs(3) = "4. FAILED: decisionFinal not persisted"
        Test_GuardarDecisionFinal_Happy = TestHelper.BuildJsonFail("field not found", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. decisionFinal persisted"

    logs(4) = "5. PASS"
    Test_GuardarDecisionFinal_Happy = TestHelper.BuildJsonOk("true", logs)
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarDecisionFinal_Happy: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarDecisionFinal_Happy = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarDecisionFinal_NotFound() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim countBefore As Long
    Dim countAfter As Long
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=0")
    logs(0) = "1. countBefore=" & countBefore & " (id=0 baseline)"

    datos.decisionFinal = "DECISION TEST"
    datos.idSolicitud = 0

    On Error Resume Next
    svc.GuardarDecisionFinal datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=0")
    If countAfter <> countBefore Then
        logs(1) = "2. FAILED: cardinality changed"
        Test_GuardarDecisionFinal_NotFound = TestHelper.BuildJsonFail("cardinality changed", logs)
        GoTo CleanExit
    End If
    If errNumber = 0 Then
        logs(1) = "2. FAILED: expected error for id=0"
        Test_GuardarDecisionFinal_NotFound = TestHelper.BuildJsonFail("Expected error for id=0", logs)
    Else
        logs(1) = "2. GuardarDecisionFinal raised Err " & errNumber & " (expected); countAfter=" & countAfter
        Test_GuardarDecisionFinal_NotFound = TestHelper.BuildJsonOk("false", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarDecisionFinal_NotFound: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarDecisionFinal_NotFound = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarDecisionFinal_Update() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim rs As DAO.Recordset
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.decisionFinal = "DECISION V1"
    datos.observacionesFinales = "OBS V1"
    datos.NombreFirmanteFinal = "FIRMANTE V1"
    datos.idSolicitud = idSol

    On Error Resume Next
    svc.GuardarDecisionFinal datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: first call raised Err " & errNumber
        Test_GuardarDecisionFinal_Update = TestHelper.BuildJsonFail("First GuardarDecisionFinal raised: " & errNumber, logs)
        GoTo CleanExit
    End If

    datos.decisionFinal = "DECISION V2"
    On Error Resume Next
    svc.GuardarDecisionFinal datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(1) = "2. FAILED: second call raised Err " & errNumber
        Test_GuardarDecisionFinal_Update = TestHelper.BuildJsonFail("Second GuardarDecisionFinal raised: " & errNumber, logs)
        GoTo CleanExit
    End If
    logs(1) = "2. Both calls succeeded"

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(2) = "3. FAILED: cardinality changed"
        Test_GuardarDecisionFinal_Update = TestHelper.BuildJsonFail("Update changed row count", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter & " (expected " & countBefore & ")"

    Set rs = db.OpenRecordset("SELECT decisionFinal FROM TbDatosCDCA WHERE idSolicitud=" & idSol, dbOpenSnapshot)
    If rs.EOF Or Nz(rs!decisionFinal, "") <> "DECISION V2" Then
        logs(3) = "4. FAILED: not updated to V2"
        Test_GuardarDecisionFinal_Update = TestHelper.BuildJsonFail("Second GuardarDecisionFinal did not update", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. updated to V2"

    logs(4) = "5. PASS"
    Test_GuardarDecisionFinal_Update = TestHelper.BuildJsonOk("true", logs)
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarDecisionFinal_Update: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarDecisionFinal_Update = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_GuardarDecisionFinal_WithoutRequiredFields() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim db As DAO.Database
    Dim idSol As Long
    Dim countBefore As Long
    Dim countAfter As Long
    Dim svc As New DatosCDCAServicio
    Dim datos As New DatosCDCA
    Dim errNumber As Long

    Set db = TestHelper.GetTestDb()
    idSol = Seed_X()
    countBefore = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    logs(0) = "1. Seeded idSol=" & idSol & " countBefore=" & countBefore

    datos.decisionFinal = ""
    datos.observacionesFinales = "OBS"
    datos.NombreFirmanteFinal = "FIRMANTE"
    datos.idSolicitud = idSol

    On Error Resume Next
    svc.GuardarDecisionFinal datos, db
    errNumber = Err.Number
    Err.Clear
    On Error GoTo EH

    countAfter = DCount("*", "TbDatosCDCA", "idSolicitud=" & idSol)
    If countAfter <> countBefore Then
        logs(1) = "2. FAILED: cardinality changed"
        Test_GuardarDecisionFinal_WithoutRequiredFields = TestHelper.BuildJsonFail("cardinality changed on validation error", logs)
        GoTo CleanExit
    End If
    If errNumber = 0 Then
        logs(1) = "2. FAILED: expected validation error"
        Test_GuardarDecisionFinal_WithoutRequiredFields = TestHelper.BuildJsonFail("Expected validation error", logs)
    Else
        logs(1) = "2. GuardarDecisionFinal raised Err " & errNumber & " (validation); countAfter=" & countAfter
        Test_GuardarDecisionFinal_WithoutRequiredFields = TestHelper.BuildJsonOk("false", logs)
    End If
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Call TestHelper.ResetTestSession
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "Test_GuardarDecisionFinal_WithoutRequiredFields: ERR " & Err.Number & " - " & Err.Description
    Test_GuardarDecisionFinal_WithoutRequiredFields = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function
