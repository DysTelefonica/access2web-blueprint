Attribute VB_Name = "Test_PCSUB"
' Attribute VB_Name = "Test_PCSUB"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Test_PCSUB.bas
' BATERÍA: PC_SUB — 84 tests — access-vba-tdd v1.9 + slices C/D/E/F (15 tests)
'
' SERVICIO: DatosPCSUBServicio.cls
'
' CICLO DE VIDA:
'   SuiteSetup  ? Configura backend TEST (una vez antes de todos)
'   SuiteTeardown ? Restaura backend original (una vez después de todos)
'   RunAll      ? Ejecuta los 84 tests y devuelve JSON resumido
'
' SLICES:
'   Slice 1: fixture graph contract (2)
'   Slice 2: core happy-path (4)
'   Slice A: completeness sad-path characterization (3)
'   Slice B1: proposal characterization (3)
'   Slice B2a: impact characterization (4)
'   Slice B2b1: impact assertion hardening (1)
'   Slice B2b2: impact validation sad-path (4)
'   Groups 5-7: existing Guardar* tests (8)
'   Groups 8-14: Es*Completa characterization tests (25)
'   Groups 15-18: Eliminar* happy/sad tests (9)
'   Slice C: GuardarAprobacionSuministrador (3)
'   Slice D: GuardarDictamenRAC + gap (4)
'   Slice E: GuardarDecisionFinal + gap (3)
'   Issue #7: ActualizarAprobacionSuministrador injected db preservation (1)
'   Slice F: PurgaTecnica + side-effects + delete (5)
'
' ==========================================================================

Private Const TEST_ID_BASE As Long = 900000
Private m_TestCounter As Long
Private m_OriginalBackend As String

' ==========================================================================
' SUITELIFECYCLE — SuiteSetup / SuiteTeardown (v1.8, exactamente 2 llamadas)
' ==========================================================================

Public Sub SuiteSetup(ByRef p_Error As String)
    Call TestHelper.SuiteSetup(p_Error)
End Sub

Public Sub SuiteTeardown(ByRef p_Error As String)
    Call TestHelper.SuiteTeardown(p_Error)
End Sub

' ==========================================================================
' HELPERS PRIVADOS — Seed / Teardown
' ==========================================================================

' --------------------------------------------------------------------------
' SeedAll: Crea base graph expediente?solicitud?tbDatosPCSUB vacío
'   Usa IDs >= TEST_ID_BASE para no colisionar con datos reales.
'   Una fila de TbExpedientes, una de tbSolicitudes, una de tbDatosPCSUB.
' --------------------------------------------------------------------------
Private Sub SeedAll()
    Dim db As DAO.Database
    Dim idExp As Long
    Dim idSol As Long
    Dim sql As String

    Set db = TestHelper.GetTestDb()

    idExp = GetNextTestId()
    idSol = GetNextTestId()

    ' tbExpedientes (Rule 1: todos los campos obligatorios)
    sql = "INSERT INTO TbExpedientes (IDExpediente, CodExp, Titulo, ObjetoContrato, ContratistaPrincipal) " & _
          "VALUES (" & idExp & ", " & TestHelper.SqlStr("EXP-" & idExp) & ", " & TestHelper.SqlStr("TEST EXP") & ", " & TestHelper.SqlStr("TEST OBJ") & ", " & TestHelper.SqlStr("Sí") & ")"
    db.Execute sql, dbFailOnError

    ' tbSolicitudes (Rule 1: todos los campos obligatorios — idEstadoInterno usa constante)
    sql = "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, " & _
          "idEstadoInterno, fechaCreacion, usuarioCreacion, fechaModificacion, usuarioModificacion, revisionCalidadEstado) " & _
          "VALUES (" & idSol & ", " & idExp & ", " & TestHelper.SqlStr("PC_SUB") & ", " & TestHelper.SqlStr("PCSUB-" & idSol) & ", " & _
          estadoPreregistro & ", Now(), " & TestHelper.SqlStr("TestPCSUB") & ", Now(), " & TestHelper.SqlStr("TestPCSUB") & ", " & TestHelper.SqlStr("PENDIENTE") & ")"
    db.Execute sql, dbFailOnError

    ' tbDatosPCSUB (vacío — requiere idSolicitud + idDatosPCSUB + refContratoInspeccionOficial)
    sql = "INSERT INTO tbDatosPCSUB (idDatosPCSUB, idSolicitud, refContratoInspeccionOficial) VALUES (" & GetNextTestId() & ", " & idSol & ", " & TestHelper.SqlStr("REF-TEST") & ")"
    db.Execute sql, dbFailOnError
End Sub

' --------------------------------------------------------------------------
' Seed_X: Crea graph para un test individual y devuelve idSolicitud
' --------------------------------------------------------------------------
Private Function Seed_X() As Long
    Dim db As DAO.Database
    Dim idExp As Long
    Dim idSol As Long
    Dim sql As String

    Set db = TestHelper.GetTestDb()

    idExp = GetNextTestId()
    idSol = GetNextTestId()

    sql = "INSERT INTO TbExpedientes (IDExpediente, CodExp, Titulo, ObjetoContrato, ContratistaPrincipal) " & _
          "VALUES (" & idExp & ", " & TestHelper.SqlStr("EXP-" & idExp) & ", " & TestHelper.SqlStr("TEST EXP") & ", " & TestHelper.SqlStr("TEST OBJ") & ", " & TestHelper.SqlStr("Sí") & ")"
    db.Execute sql, dbFailOnError

    sql = "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, " & _
          "idEstadoInterno, fechaCreacion, usuarioCreacion, fechaModificacion, usuarioModificacion, revisionCalidadEstado) " & _
          "VALUES (" & idSol & ", " & idExp & ", " & TestHelper.SqlStr("PC_SUB") & ", " & TestHelper.SqlStr("PCSUB-" & idSol) & ", " & _
          estadoPreregistro & ", Now(), " & TestHelper.SqlStr("TestPCSUB") & ", Now(), " & TestHelper.SqlStr("TestPCSUB") & ", " & TestHelper.SqlStr("PENDIENTE") & ")"
    db.Execute sql, dbFailOnError

    sql = "INSERT INTO tbDatosPCSUB (idDatosPCSUB, idSolicitud, refContratoInspeccionOficial) VALUES (" & GetNextTestId() & ", " & idSol & ", " & TestHelper.SqlStr("REF-TEST") & ")"
    db.Execute sql, dbFailOnError

    Seed_X = idSol
End Function

' --------------------------------------------------------------------------
' TeardownAll: Limpia en orden FK inverso (v1.9)
' --------------------------------------------------------------------------
Public Sub TeardownAll()
    Dim db As DAO.Database
    Dim sql As String

    Set db = TestHelper.GetTestDb()
    On Error Resume Next

    sql = "DELETE FROM tbDatosPCSUB WHERE idSolicitud >= " & TEST_ID_BASE
    db.Execute sql, dbFailOnError

    sql = "DELETE FROM tbSolicitudes WHERE idSolicitud >= " & TEST_ID_BASE
    db.Execute sql, dbFailOnError

    sql = "DELETE FROM TbExpedientes WHERE IDExpediente >= " & TEST_ID_BASE
    db.Execute sql, dbFailOnError
End Sub

' --------------------------------------------------------------------------
' GetNextTestId: Contador incremental para IDs de test (TEST_ID_BASE + secuencia)
' --------------------------------------------------------------------------
Private Function GetNextTestId() As Long
    m_TestCounter = m_TestCounter + 1
    GetNextTestId = TEST_ID_BASE + m_TestCounter
End Function

' --------------------------------------------------------------------------
' SeedCompletePCSUBFixture: Crea graph completo con todos los field groups
'   Used by slices C/D/E/F for side-effect and characterization tests
' --------------------------------------------------------------------------
Private Function SeedCompletePCSUBFixture() As Long
    Dim db As DAO.Database
    Dim idExp As Long
    Dim idSol As Long
    Dim sql As String

    Set db = TestHelper.GetTestDb()

    idExp = GetNextTestId()
    idSol = GetNextTestId()

    ' tbExpedientes
    sql = "INSERT INTO TbExpedientes (IDExpediente, CodExp, Titulo, ObjetoContrato, ContratistaPrincipal) " & _
          "VALUES (" & idExp & ", " & TestHelper.SqlStr("EXP-" & idExp) & ", " & TestHelper.SqlStr("TEST EXP") & ", " & TestHelper.SqlStr("TEST OBJ") & ", " & TestHelper.SqlStr("Sí") & ")"
    db.Execute sql, dbFailOnError

    ' tbSolicitudes
    sql = "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, " & _
          "idEstadoInterno, fechaCreacion, usuarioCreacion, fechaModificacion, usuarioModificacion, revisionCalidadEstado) " & _
          "VALUES (" & idSol & ", " & idExp & ", " & TestHelper.SqlStr("PC_SUB") & ", " & TestHelper.SqlStr("PCSUB-" & idSol) & ", " & _
          estadoPreregistro & ", Now(), " & TestHelper.SqlStr("TestPCSUB") & ", Now(), " & TestHelper.SqlStr("TestPCSUB") & ", " & TestHelper.SqlStr("PENDIENTE") & ")"
    db.Execute sql, dbFailOnError

    ' tbDatosPCSUB con todos los field groups completos
    sql = "INSERT INTO tbDatosPCSUB (idDatosPCSUB, idSolicitud, " & _
          "refContratoInspeccionOficial, refSubSuministrador, denominacionContrato, SubsuministradorNombreDir, objetoContrato, " & _
          "descripcionMaterialAfectado, numPlanoEspecificacion, descripcionPropuestaCambio, descripcionPropuestaCambioCont, " & _
          "motivoCorregirDeficiencias, motivoMejorarCapacidad, motivoAumentarNacionalizacion, motivoMejorarSeguridad, " & _
          "motivoMejorarFiabilidad, motivoMejorarCosteEficacia, motivoOtros, motivoOtrosDetalle, " & _
          "incidenciaCoste, incidenciaPlazo, incidenciaSeguridad, incidenciaFiabilidad, incidenciaMantenibilidad, " & _
          "incidenciaIntercambiabilidad, incidenciaVidaUtilAlmacen, incidenciaFuncionamientoFuncion, " & _
          "impactoClasificacion, CambioAfectaAMaterial, " & _
          "firmaOficinaTecnicaSubSuministradorNombre, firmaRepSubSuministradorNombre, " & _
          "racCodigo, observacionesRAC, racNombre, racDecision, racRechazoMotivos, " & _
          "observacionesRACDelegador, racNombreDelegador, " & _
          "obsAprobacionAutoridadDiseno, NombreAutoridadDiseno, " & _
          "decisionFinal, obsDecisionFinal, NombreFirmanteFinal) " & _
          "VALUES (" & GetNextTestId() & ", " & idSol & ", " & _
          TestHelper.SqlStr("REF-C-001") & ", " & TestHelper.SqlStr("REF-SUB-001") & ", " & TestHelper.SqlStr("DENOM TEST") & ", " & TestHelper.SqlStr("DIR TEST") & ", " & TestHelper.SqlStr("OBJ TEST") & ", " & _
          TestHelper.SqlStr("MATERIAL AFECTADO") & ", " & TestHelper.SqlStr("PLANO-001") & ", " & TestHelper.SqlStr("PROPUESTA CAMBIO") & ", " & TestHelper.SqlStr("CONTINUACION") & ", " & _
          "True, True, False, True, False, True, False, " & TestHelper.SqlStr("Detalle otros") & ", " & _
          TestHelper.SqlStr("AUMENTARÁ") & ", " & TestHelper.SqlStr("NO VARIARÁ") & ", True, False, False, False, False, False, " & _
          TestHelper.SqlStr("MAYOR") & ", " & TestHelper.SqlStr("SÍ") & ", " & _
          TestHelper.SqlStr("FIRMA OFICINA TECNICA") & ", " & TestHelper.SqlStr("FIRMA REP SUB") & ", " & _
          TestHelper.SqlStr("RAC-2024-001") & ", " & TestHelper.SqlStr("OBS RAC") & ", " & TestHelper.SqlStr("NOMBRE RAC") & ", " & TestHelper.SqlStr("APROBADO") & ", " & TestHelper.SqlStr("MOTIVOS RECHAZO") & ", " & _
          TestHelper.SqlStr("OBS DELEGADOR") & ", " & TestHelper.SqlStr("NOMBRE DELEGADOR") & ", " & _
          TestHelper.SqlStr("OBS AUTORIDAD") & ", " & TestHelper.SqlStr("NOMBRE AUTORIDAD") & ", " & _
          TestHelper.SqlStr("DECISIÓN FINAL") & ", " & TestHelper.SqlStr("OBS DECISION") & ", " & TestHelper.SqlStr("NOMBRE FIRMANTE") & ")"
    db.Execute sql, dbFailOnError

    SeedCompletePCSUBFixture = idSol
End Function

' --------------------------------------------------------------------------
' SeedPCSUBTechnicalFixture: Crea graph con campos Tecnica completos
'   Used by PurgaTecnica side-effect tests
' --------------------------------------------------------------------------
Private Function SeedPCSUBTechnicalFixture() As Long
    Dim db As DAO.Database
    Dim idExp As Long
    Dim idSol As Long
    Dim sql As String

    Set db = TestHelper.GetTestDb()

    idExp = GetNextTestId()
    idSol = GetNextTestId()

    ' tbExpedientes
    sql = "INSERT INTO TbExpedientes (IDExpediente, CodExp, Titulo, ObjetoContrato, ContratistaPrincipal) " & _
          "VALUES (" & idExp & ", " & TestHelper.SqlStr("EXP-" & idExp) & ", " & TestHelper.SqlStr("TEST EXP") & ", " & TestHelper.SqlStr("TEST OBJ") & ", " & TestHelper.SqlStr("Sí") & ")"
    db.Execute sql, dbFailOnError

    ' tbSolicitudes
    sql = "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, " & _
          "idEstadoInterno, fechaCreacion, usuarioCreacion, fechaModificacion, usuarioModificacion, revisionCalidadEstado) " & _
          "VALUES (" & idSol & ", " & idExp & ", " & TestHelper.SqlStr("PC_SUB") & ", " & TestHelper.SqlStr("PCSUB-" & idSol) & ", " & _
          estadoPreregistro & ", Now(), " & TestHelper.SqlStr("TestPCSUB") & ", Now(), " & TestHelper.SqlStr("TestPCSUB") & ", " & TestHelper.SqlStr("PENDIENTE") & ")"
    db.Execute sql, dbFailOnError

    ' tbDatosPCSUB con Propuesta + Impacto completos
    sql = "INSERT INTO tbDatosPCSUB (idDatosPCSUB, idSolicitud, refContratoInspeccionOficial, " & _
          "descripcionMaterialAfectado, numPlanoEspecificacion, descripcionPropuestaCambio, descripcionPropuestaCambioCont, " & _
          "motivoCorregirDeficiencias, motivoMejorarCapacidad, motivoAumentarNacionalizacion, motivoMejorarSeguridad, " & _
          "motivoMejorarFiabilidad, motivoMejorarCosteEficacia, motivoOtros, motivoOtrosDetalle, " & _
          "incidenciaCoste, incidenciaPlazo, incidenciaSeguridad, incidenciaFiabilidad, incidenciaMantenibilidad, " & _
          "incidenciaIntercambiabilidad, incidenciaVidaUtilAlmacen, incidenciaFuncionamientoFuncion, " & _
          "impactoClasificacion, CambioAfectaAMaterial) " & _
          "VALUES (" & GetNextTestId() & ", " & idSol & ", " & TestHelper.SqlStr("REF-TECH-001") & ", " & _
          TestHelper.SqlStr("MATERIAL AFECTADO") & ", " & TestHelper.SqlStr("PLANO-001") & ", " & TestHelper.SqlStr("PROPUESTA CAMBIO") & ", " & TestHelper.SqlStr("CONTINUACION") & ", " & _
          "True, True, False, True, False, True, False, " & TestHelper.SqlStr("Detalle otros") & ", " & _
          TestHelper.SqlStr("AUMENTARÁ") & ", " & TestHelper.SqlStr("NO VARIARÁ") & ", True, False, False, False, False, False, " & _
          TestHelper.SqlStr("MAYOR") & ", " & TestHelper.SqlStr("SÍ") & ")"
    db.Execute sql, dbFailOnError

    SeedPCSUBTechnicalFixture = idSol
End Function

' --------------------------------------------------------------------------
' Helper: AprobacionFieldsMatch — verifica que los campos de firma son los esperados
' --------------------------------------------------------------------------
Private Function AprobacionFieldsMatch(idSol As Long, ByVal expS1 As String, ByVal expS2 As String) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    Set db = TestHelper.GetTestDb()
    sql = "SELECT firmaOficinaTecnicaSubSuministradorNombre, firmaRepSubSuministradorNombre FROM tbDatosPCSUB WHERE idSolicitud = " & idSol
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        AprobacionFieldsMatch = False
    Else
        AprobacionFieldsMatch = (Nz(rs!firmaOficinaTecnicaSubSuministradorNombre, "") = expS1 And _
                                 Nz(rs!firmaRepSubSuministradorNombre, "") = expS2)
    End If

    rs.Close
    Set rs = Nothing
End Function

' --------------------------------------------------------------------------
' Helper: RacFieldsMatch — verifica que los campos RAC son los esperados
' --------------------------------------------------------------------------
Private Function RacFieldsMatch(idSol As Long, ByVal expCodigo As String, ByVal expDecision As String) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    Set db = TestHelper.GetTestDb()
    sql = "SELECT racCodigo, racDecision FROM tbDatosPCSUB WHERE idSolicitud = " & idSol
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        RacFieldsMatch = False
    Else
        RacFieldsMatch = (Nz(rs!racCodigo, "") = expCodigo And _
                          Nz(rs!racDecision, "") = expDecision)
    End If

    rs.Close
    Set rs = Nothing
End Function

' --------------------------------------------------------------------------
' Helper: DecisionFieldsMatch — verifica que los campos Decision Final son los esperados
' --------------------------------------------------------------------------
Private Function DecisionFieldsMatch(idSol As Long, ByVal expDecision As String, ByVal expFirmante As String) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    Set db = TestHelper.GetTestDb()
    sql = "SELECT decisionFinal, NombreFirmanteFinal FROM tbDatosPCSUB WHERE idSolicitud = " & idSol
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        DecisionFieldsMatch = False
    Else
        DecisionFieldsMatch = (Nz(rs!decisionFinal, "") = expDecision And _
                               Nz(rs!NombreFirmanteFinal, "") = expFirmante)
    End If

    rs.Close
    Set rs = Nothing
End Function

' --------------------------------------------------------------------------
' Helper: TecnicaFieldsCleared — verifica que propuesta+impacto fueron limpiados
' --------------------------------------------------------------------------
Private Function TecnicaFieldsCleared(idSol As Long) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    Set db = TestHelper.GetTestDb()
    sql = "SELECT descripcionMaterialAfectado, descripcionPropuestaCambio, incidenciaCoste, impactoClasificacion " & _
          "FROM tbDatosPCSUB WHERE idSolicitud = " & idSol
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        TecnicaFieldsCleared = False
    Else
        TecnicaFieldsCleared = (Nz(rs!descripcionMaterialAfectado, "") = "" And _
                                Nz(rs!descripcionPropuestaCambio, "") = "" And _
                                Nz(rs!incidenciaCoste, "") = "" And _
                                Nz(rs!impactoClasificacion, "") = "")
    End If

    rs.Close
    Set rs = Nothing
End Function

' --------------------------------------------------------------------------
' Helper: AprobacionFieldsCleared — verifica que los campos de firma fueron limpiados
' --------------------------------------------------------------------------
Private Function AprobacionFieldsCleared(idSol As Long) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    Set db = TestHelper.GetTestDb()
    sql = "SELECT firmaOficinaTecnicaSubSuministradorNombre, firmaRepSubSuministradorNombre " & _
          "FROM tbDatosPCSUB WHERE idSolicitud = " & idSol
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        AprobacionFieldsCleared = False
    Else
        AprobacionFieldsCleared = (Nz(rs!firmaOficinaTecnicaSubSuministradorNombre, "") = "" And _
                                   Nz(rs!firmaRepSubSuministradorNombre, "") = "")
    End If

    rs.Close
    Set rs = Nothing
End Function

' --------------------------------------------------------------------------
' Helper: RacFieldsCleared — verifica que los campos RAC fueron limpiados
' --------------------------------------------------------------------------
Private Function RacFieldsCleared(idSol As Long) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    Set db = TestHelper.GetTestDb()
    sql = "SELECT racCodigo, racDecision, racNombre FROM tbDatosPCSUB WHERE idSolicitud = " & idSol
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        RacFieldsCleared = False
    Else
        RacFieldsCleared = (Nz(rs!racCodigo, "") = "" And _
                            Nz(rs!racDecision, "") = "" And _
                            Nz(rs!racNombre, "") = "")
    End If

    rs.Close
    Set rs = Nothing
End Function

' --------------------------------------------------------------------------
' Helper: DecisionFieldsCleared — verifica que los campos Decision Final fueron limpiados
' --------------------------------------------------------------------------
Private Function DecisionFieldsCleared(idSol As Long) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    Set db = TestHelper.GetTestDb()
    sql = "SELECT decisionFinal, NombreFirmanteFinal, obsDecisionFinal " & _
          "FROM tbDatosPCSUB WHERE idSolicitud = " & idSol
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        DecisionFieldsCleared = False
    Else
        DecisionFieldsCleared = (Nz(rs!decisionFinal, "") = "" And _
                                  Nz(rs!NombreFirmanteFinal, "") = "" And _
                                  Nz(rs!obsDecisionFinal, "") = "")
    End If

    rs.Close
    Set rs = Nothing
End Function

' --------------------------------------------------------------------------
' Helper: CountPCSUBRows — cuenta filas de tbDatosPCSUB para un idSolicitud
' --------------------------------------------------------------------------
Private Function CountPCSUBRows(idSol As Long) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    Set db = TestHelper.GetTestDb()
    sql = "SELECT COUNT(*) AS cnt FROM tbDatosPCSUB WHERE idSolicitud = " & idSol
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        CountPCSUBRows = 0
    Else
        CountPCSUBRows = rs!cnt
    End If

    rs.Close
    Set rs = Nothing
End Function

' ==========================================================================
' RUNALL — Ejecutor principal
' ==========================================================================

Public Function Test_PCSUB_RunAll() As String
    Test_PCSUB_RunAll = RunAll()
End Function

Public Function RunAll() As String
    On Error GoTo EH

    Dim passed As Long
    Dim failed As Long
    Dim resultsJson As String
    Dim resultJson As String
    Dim logs(0 To 3) As String

    resultsJson = "["

    resultJson = Test_PCSUB_GuardarAprobacionSuministrador_PersistsBothSigners(): RecordRunAllResult resultJson, passed, failed, resultsJson
    resultJson = Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing(): RecordRunAllResult resultJson, passed, failed, resultsJson
    resultJson = Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing(): RecordRunAllResult resultJson, passed, failed, resultsJson
    resultJson = Test_PCSUB_GuardarDictamenRAC_PersistsValidRAC(): RecordRunAllResult resultJson, passed, failed, resultsJson
    resultJson = Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired(): RecordRunAllResult resultJson, passed, failed, resultsJson
    resultJson = Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected(): RecordRunAllResult resultJson, passed, failed, resultsJson
    resultJson = Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse(): RecordRunAllResult resultJson, passed, failed, resultsJson
    resultJson = Test_PCSUB_GuardarDecisionFinal_PersistsValidDecision(): RecordRunAllResult resultJson, passed, failed, resultsJson
    resultJson = Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty(): RecordRunAllResult resultJson, passed, failed, resultsJson
    resultJson = Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse(): RecordRunAllResult resultJson, passed, failed, resultsJson
    resultJson = Test_PCSUB_PurgaTecnica_ClearsAllTecnicaFields(): RecordRunAllResult resultJson, passed, failed, resultsJson
    resultJson = Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields(): RecordRunAllResult resultJson, passed, failed, resultsJson
    resultJson = Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields(): RecordRunAllResult resultJson, passed, failed, resultsJson
    resultJson = Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields(): RecordRunAllResult resultJson, passed, failed, resultsJson
    resultJson = Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow(): RecordRunAllResult resultJson, passed, failed, resultsJson

    resultsJson = resultsJson & "]"
    logs(0) = "Test_PCSUB_RunAll: public smoke subset completed"
    logs(1) = "passed=" & passed
    logs(2) = "failed=" & failed
    logs(3) = "results=" & resultsJson

    If failed > 0 Then
        RunAll = TestHelper.BuildJsonFail("PCSUB smoke failed: " & failed, logs)
    Else
        RunAll = TestHelper.BuildJsonOk("passed=" & passed, logs)
    End If
    Exit Function

EH:
    logs(0) = "Test_PCSUB_RunAll: ERR " & Err.Number & " - " & Err.description
    RunAll = TestHelper.BuildJsonFail(Err.description, logs)
End Function

Private Sub RecordRunAllResult(ByVal resultJson As String, ByRef passed As Long, ByRef failed As Long, ByRef resultsJson As String)
    If Len(resultsJson) > 1 Then
        resultsJson = resultsJson & ","
    End If
    resultsJson = resultsJson & resultJson

    If Left$(resultJson, 10) = "{""ok"":true" Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
End Sub

' ==========================================================================
' SLICE C — GuardarAprobacionSuministrador (3 tests)
' ==========================================================================

Public Function Test_PCSUB_GuardarAprobacionSuministrador_PersistsBothSigners() As String
    On Error GoTo EH
    Dim logs() As String: logs = NewLogsArray(6)
    Dim runError As String
    Dim idSol As Long
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim db As DAO.Database

    TestHelper.ForceLocalBackend runError
    If runError <> "" Then
        logs(0) = runError
        Test_PCSUB_GuardarAprobacionSuministrador_PersistsBothSigners = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Call TeardownAll

    idSol = SeedCompletePCSUBFixture()
    Set db = TestHelper.GetTestDb()

    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.firmaOficinaTecnicaSubSuministradorNombre = "Ing. Juan Pérez"
    pcsub.firmaRepSubSuministradorNombre = "Sra. María López"

    svc.GuardarAprobacionSuministrador pcsub, db

    If Not AprobacionFieldsMatch(idSol, "Ing. Juan Pérez", "Sra. María López") Then
        logs(4) = "Test_PCSUB_GuardarAprobacionSuministrador_PersistsBothSigners: FAILED - fields not persisted"
        logs(5) = "Test_PCSUB_GuardarAprobacionSuministrador_PersistsBothSigners: complete"
        Test_PCSUB_GuardarAprobacionSuministrador_PersistsBothSigners = TestHelper.BuildJsonFail("fields not persisted", logs)
        GoTo CleanExit
    End If

    logs(0) = "Test_PCSUB_GuardarAprobacionSuministrador_PersistsBothSigners: seed complete fixture C_BOTH"
    logs(1) = "Test_PCSUB_GuardarAprobacionSuministrador_PersistsBothSigners: GuardarAprobacionSuministrador succeeded"
    logs(2) = "Test_PCSUB_GuardarAprobacionSuministrador_PersistsBothSigners: both signer names persisted correctly"
    logs(3) = "Test_PCSUB_GuardarAprobacionSuministrador_PersistsBothSigners: AprobacionFieldsMatch verified"
    logs(4) = "Test_PCSUB_GuardarAprobacionSuministrador_PersistsBothSigners: complete"
    logs(5) = "Test_PCSUB_GuardarAprobacionSuministrador_PersistsBothSigners: PASS"
    Test_PCSUB_GuardarAprobacionSuministrador_PersistsBothSigners = TestHelper.BuildJsonOk("true", logs)
CleanExit:
    On Error Resume Next
    Call TeardownAll
    Call TestHelper.ResetTestSession
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_GuardarAprobacionSuministrador_PersistsBothSigners: ERR " & Err.Number & " - " & Err.description
    Test_PCSUB_GuardarAprobacionSuministrador_PersistsBothSigners = TestHelper.BuildJsonFail(Err.description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing() As String
    On Error GoTo EH
    Dim logs() As String: logs = NewLogsArray(6)
    Dim runError As String
    Dim idSol As Long
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim db As DAO.Database
    Dim errCaught As Boolean
    Dim errNumber As Long

    TestHelper.ForceLocalBackend runError
    If runError <> "" Then
        logs(0) = runError
        Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Call TeardownAll

    idSol = SeedCompletePCSUBFixture()
    Set db = TestHelper.GetTestDb()

    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.firmaOficinaTecnicaSubSuministradorNombre = ""  ' Missing!
    pcsub.firmaRepSubSuministradorNombre = "Sra. María López"

    On Error Resume Next
    errCaught = False
    svc.GuardarAprobacionSuministrador pcsub, db
    errNumber = Err.Number
    If errNumber <> 0 Then errCaught = True
    Err.Clear
    On Error GoTo EH

    If Not errCaught Or errNumber <> 513 Then
        logs(4) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing: FAILED - expected error 513"
        logs(5) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing: complete"
        Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing = TestHelper.BuildJsonFail("expected error 513 not raised", logs)
        GoTo CleanExit
    End If

    If Not AprobacionFieldsMatch(idSol, "FIRMA OFICINA TECNICA", "FIRMA REP SUB") Then
        logs(4) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing: FAILED - validation error changed persisted fields"
        logs(5) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing: complete"
        Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing = TestHelper.BuildJsonFail("validation error changed persisted fields", logs)
        GoTo CleanExit
    End If

    logs(0) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing: seed complete fixture C_S1MISS"
    logs(1) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing: signer1 empty"
    logs(2) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing: GuardarAprobacionSuministrador raised 513 as expected"
    logs(3) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing: error caught and validated"
    logs(4) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing: complete"
    logs(5) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing: PASS"
    Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing = TestHelper.BuildJsonOk("true", logs)
CleanExit:
    On Error Resume Next
    Call TeardownAll
    Call TestHelper.ResetTestSession
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing: ERR " & Err.Number & " - " & Err.description
    Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner1Missing = TestHelper.BuildJsonFail(Err.description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing() As String
    On Error GoTo EH
    Dim logs() As String: logs = NewLogsArray(6)
    Dim runError As String
    Dim idSol As Long
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim db As DAO.Database
    Dim errCaught As Boolean
    Dim errNumber As Long

    TestHelper.ForceLocalBackend runError
    If runError <> "" Then
        logs(0) = runError
        Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Call TeardownAll

    idSol = SeedCompletePCSUBFixture()
    Set db = TestHelper.GetTestDb()

    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.firmaOficinaTecnicaSubSuministradorNombre = "Ing. Juan Pérez"
    pcsub.firmaRepSubSuministradorNombre = ""  ' Missing!

    On Error Resume Next
    errCaught = False
    svc.GuardarAprobacionSuministrador pcsub, db
    errNumber = Err.Number
    If errNumber <> 0 Then errCaught = True
    Err.Clear
    On Error GoTo EH

    If Not errCaught Or errNumber <> 513 Then
        logs(4) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing: FAILED - expected error 513"
        logs(5) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing: complete"
        Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing = TestHelper.BuildJsonFail("expected error 513 not raised", logs)
        GoTo CleanExit
    End If

    If Not AprobacionFieldsMatch(idSol, "FIRMA OFICINA TECNICA", "FIRMA REP SUB") Then
        logs(4) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing: FAILED - validation error changed persisted fields"
        logs(5) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing: complete"
        Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing = TestHelper.BuildJsonFail("validation error changed persisted fields", logs)
        GoTo CleanExit
    End If

    logs(0) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing: seed complete fixture C_S2MISS"
    logs(1) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing: signer2 empty"
    logs(2) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing: GuardarAprobacionSuministrador raised 513 as expected"
    logs(3) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing: error caught and validated"
    logs(4) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing: complete"
    logs(5) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing: PASS"
    Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing = TestHelper.BuildJsonOk("true", logs)
CleanExit:
    On Error Resume Next
    Call TeardownAll
    Call TestHelper.ResetTestSession
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing: ERR " & Err.Number & " - " & Err.description
    Test_PCSUB_GuardarAprobacionSuministrador_FailsWhenSigner2Missing = TestHelper.BuildJsonFail(Err.description, logs)
    Resume CleanExit
End Function

' (Issue #7 test moved to Test_Issue7_DbInjection.bas — surgical commit
'  for GH #7 closeout. Other Slice A-F tests follow below.)

' ==========================================================================
' SLICE D — GuardarDictamenRAC (4 tests)
' ==========================================================================

Public Function Test_PCSUB_GuardarDictamenRAC_PersistsValidRAC() As String
    On Error GoTo EH
    Dim logs() As String: logs = NewLogsArray(6)
    Dim runError As String
    Dim idSol As Long
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim db As DAO.Database

    TestHelper.ForceLocalBackend runError
    If runError <> "" Then
        logs(0) = runError
        Test_PCSUB_GuardarDictamenRAC_PersistsValidRAC = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Call TeardownAll

    idSol = SeedCompletePCSUBFixture()
    Set db = TestHelper.GetTestDb()

    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.racCodigo = "RAC-2024-NEW"
    pcsub.racDecision = "APROBADO"
    pcsub.racNombre = "Dictamen Nuevo"

    svc.GuardarDictamenRAC pcsub, db

    If Not RacFieldsMatch(idSol, "RAC-2024-NEW", "APROBADO") Then
        logs(4) = "Test_PCSUB_GuardarDictamenRAC_PersistsValidRAC: FAILED - RAC fields not persisted"
        logs(5) = "Test_PCSUB_GuardarDictamenRAC_PersistsValidRAC: complete"
        Test_PCSUB_GuardarDictamenRAC_PersistsValidRAC = TestHelper.BuildJsonFail("RAC fields not persisted", logs)
        GoTo CleanExit
    End If

    logs(0) = "Test_PCSUB_GuardarDictamenRAC_PersistsValidRAC: seed complete fixture D1RAC"
    logs(1) = "Test_PCSUB_GuardarDictamenRAC_PersistsValidRAC: racCodigo + racDecision set"
    logs(2) = "Test_PCSUB_GuardarDictamenRAC_PersistsValidRAC: GuardarDictamenRAC succeeded"
    logs(3) = "Test_PCSUB_GuardarDictamenRAC_PersistsValidRAC: RacFieldsMatch verified"
    logs(4) = "Test_PCSUB_GuardarDictamenRAC_PersistsValidRAC: complete"
    logs(5) = "Test_PCSUB_GuardarDictamenRAC_PersistsValidRAC: PASS"
    Test_PCSUB_GuardarDictamenRAC_PersistsValidRAC = TestHelper.BuildJsonOk("true", logs)
CleanExit:
    On Error Resume Next
    Call TeardownAll
    Call TestHelper.ResetTestSession
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_GuardarDictamenRAC_PersistsValidRAC: ERR " & Err.Number & " - " & Err.description
    Test_PCSUB_GuardarDictamenRAC_PersistsValidRAC = TestHelper.BuildJsonFail(Err.description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired() As String
    On Error GoTo EH
    Dim logs() As String: logs = NewLogsArray(6)
    Dim runError As String
    Dim idSol As Long
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim db As DAO.Database
    Dim errNumber As Long
    Dim errDescription As String

    TestHelper.ForceLocalBackend runError
    If runError <> "" Then
        logs(0) = runError
        Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Call TeardownAll

    idSol = SeedCompletePCSUBFixture()
    Set db = TestHelper.GetTestDb()

    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.racCodigo = ""  ' Empty is OK for RECHAZADO
    pcsub.racDecision = "RECHAZADO"
    pcsub.racNombre = "Dictamen Rechazo"

    On Error Resume Next
    svc.GuardarDictamenRAC pcsub, db
    errNumber = Err.Number
    errDescription = Err.description
    Err.Clear
    On Error GoTo EH

    If errNumber <> 0 Then
        logs(4) = "Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired: FAILED - unexpected error"
        logs(5) = "Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired: complete"
        Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired = TestHelper.BuildJsonFail("unexpected error " & errNumber & ": " & errDescription, logs)
        GoTo CleanExit
    End If

    If Not RacFieldsMatch(idSol, "", "RECHAZADO") Then
        logs(4) = "Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired: FAILED - RECHAZADO fields not persisted"
        logs(5) = "Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired: complete"
        Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired = TestHelper.BuildJsonFail("RECHAZADO fields not persisted", logs)
        GoTo CleanExit
    End If

    logs(0) = "Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired: seed complete fixture D2REJ"
    logs(1) = "Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired: racDecision=RECHAZADO, racCodigo empty"
    logs(2) = "Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired: GuardarDictamenRAC succeeded (no error)"
    logs(3) = "Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired: confirmed RECHAZADO+empty code persisted"
    logs(4) = "Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired: complete"
    logs(5) = "Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired: PASS"
    Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired = TestHelper.BuildJsonOk("true", logs)
CleanExit:
    On Error Resume Next
    Call TeardownAll
    Call TestHelper.ResetTestSession
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired: ERR " & Err.Number & " - " & Err.description
    Test_PCSUB_GuardarDictamenRAC_RejectedDecision_NocodeRequired = TestHelper.BuildJsonFail(Err.description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected() As String
    On Error GoTo EH
    Dim logs() As String: logs = NewLogsArray(6)
    Dim runError As String
    Dim idSol As Long
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim db As DAO.Database
    Dim errCaught As Boolean
    Dim errNumber As Long
    Dim errDescription As String
    Dim expectedDescription As String

    TestHelper.ForceLocalBackend runError
    If runError <> "" Then
        logs(0) = runError
        Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Call TeardownAll

    idSol = SeedCompletePCSUBFixture()
    Set db = TestHelper.GetTestDb()

    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.racCodigo = ""  ' Missing!
    pcsub.racDecision = "APROBADO"  ' Not RECHAZADO ? code required
    pcsub.racNombre = "Dictamen"

    On Error Resume Next
    errCaught = False
    svc.GuardarDictamenRAC pcsub, db
    errNumber = Err.Number
    errDescription = Err.description
    If errNumber <> 0 Then errCaught = True
    Err.Clear
    On Error GoTo EH
    expectedDescription = "Error #513: Código RAC obligatorio." & vbNewLine & "Pila de Llamadas:" & vbNewLine & "   en DatosPCSUBServicio.GuardarDictamenRAC"

    If Not errCaught Or errNumber <> 513 Or errDescription <> expectedDescription Then
        logs(4) = "Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected: FAILED - expected error 513"
        logs(5) = "Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected: complete"
        Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected = TestHelper.BuildJsonFail("expected exact error 513 message not raised", logs)
        GoTo CleanExit
    End If

    If Not RacFieldsMatch(idSol, "RAC-2024-001", "APROBADO") Then
        logs(4) = "Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected: FAILED - validation error changed persisted fields"
        logs(5) = "Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected: complete"
        Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected = TestHelper.BuildJsonFail("validation error changed persisted fields", logs)
        GoTo CleanExit
    End If

    logs(0) = "Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected: seed complete fixture D3NOCODE"
    logs(1) = "Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected: racCodigo empty, racDecision=APROBADO"
    logs(2) = "Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected: GuardarDictamenRAC raised exact 513 message as expected"
    logs(3) = "Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected: validation error left persisted RAC fields unchanged"
    logs(4) = "Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected: complete"
    logs(5) = "Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected: PASS"
    Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected = TestHelper.BuildJsonOk("true", logs)
CleanExit:
    On Error Resume Next
    Call TeardownAll
    Call TestHelper.ResetTestSession
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected: ERR " & Err.Number & " - " & Err.description
    Test_PCSUB_GuardarDictamenRAC_MissingCode_NonRejected = TestHelper.BuildJsonFail(Err.description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse() As String
    On Error GoTo EH
    Dim logs() As String: logs = NewLogsArray(6)
    Dim runError As String
    Dim idSol As Long
    Dim svc As New DatosPCSUBServicio
    Dim result As Boolean
    Dim db As DAO.Database

    TestHelper.ForceLocalBackend runError
    If runError <> "" Then
        logs(0) = runError
        Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Call TeardownAll

    idSol = SeedCompletePCSUBFixture()
    Set db = TestHelper.GetTestDb()

    ' Pre-condition: racCodigo and racNombre populated, but racDecision blank
    db.Execute "UPDATE tbDatosPCSUB SET racDecision = '' WHERE idSolicitud = " & idSol, dbFailOnError
    result = svc.EsDictamenRACCompleto(idSol, db)

    If result Then
        logs(4) = "Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse: FAILED - expected fix to return False when racDecision blank"
        logs(5) = "Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse: complete"
        Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse = TestHelper.BuildJsonFail("expected fix to return False", logs)
        GoTo CleanExit
    End If

    ' FIX verified: EsDictamenRACCompleto now returns False when racDecision is blank
    logs(0) = "Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse: seed complete fixture D4GAP"
    logs(1) = "Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse: racCodigo + racNombre populated, racDecision blank"
    logs(2) = "Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse: EsDictamenRACCompleto returned " & result
    logs(3) = "Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse: FIX verified - returns false when decision blank"
    logs(4) = "Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse: complete"
    logs(5) = "Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse: PASS"
    Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse = TestHelper.BuildJsonOk(CStr(result), logs)
CleanExit:
    On Error Resume Next
    Call TeardownAll
    Call TestHelper.ResetTestSession
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse: ERR " & Err.Number & " - " & Err.description
    Test_PCSUB_EsDictamenRACCompleto_BlankDecision_ReturnsFalse = TestHelper.BuildJsonFail(Err.description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_EsAprobacionSuministradorCompleta_BlankDecisionFinal_ReturnsFalse() As String
    On Error GoTo EH
    Dim logs() As String: logs = NewLogsArray(6)
    Dim runError As String
    Dim idSol As Long
    Dim svc As New DatosPCSUBServicio
    Dim result As Boolean
    Dim db As DAO.Database

    TestHelper.ForceLocalBackend runError
    If runError <> "" Then
        logs(0) = runError
        Test_PCSUB_EsAprobacionSuministradorCompleta_BlankDecisionFinal_ReturnsFalse = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Call TeardownAll

    idSol = SeedCompletePCSUBFixture()
    Set db = TestHelper.GetTestDb()

    ' Pre-condition: both firma fields populated, but decisionFinal blank
    db.Execute "UPDATE tbDatosPCSUB SET decisionFinal = '' WHERE idSolicitud = " & idSol, dbFailOnError
    result = svc.EsAprobacionSuministradorCompleta(idSol, db)

    If result Then
        logs(4) = "Test_PCSUB_EsAprobacionSuministradorCompleta_BlankDecisionFinal_ReturnsFalse: FAILED - expected fix to return False when decisionFinal blank"
        logs(5) = "Test_PCSUB_EsAprobacionSuministradorCompleta_BlankDecisionFinal_ReturnsFalse: complete"
        Test_PCSUB_EsAprobacionSuministradorCompleta_BlankDecisionFinal_ReturnsFalse = TestHelper.BuildJsonFail("expected fix to return False", logs)
        GoTo CleanExit
    End If

    ' FIX verified: EsAprobacionSuministradorCompleta now returns False when decisionFinal is blank
    logs(0) = "Test_PCSUB_EsAprobacionSuministradorCompleta_BlankDecisionFinal_ReturnsFalse: seed complete fixture C_DECBLANK"
    logs(1) = "Test_PCSUB_EsAprobacionSuministradorCompleta_BlankDecisionFinal_ReturnsFalse: both firma fields populated, decisionFinal blank"
    logs(2) = "Test_PCSUB_EsAprobacionSuministradorCompleta_BlankDecisionFinal_ReturnsFalse: EsAprobacionSuministradorCompleta returned " & result
    logs(3) = "Test_PCSUB_EsAprobacionSuministradorCompleta_BlankDecisionFinal_ReturnsFalse: FIX verified - returns false when decisionFinal blank"
    logs(4) = "Test_PCSUB_EsAprobacionSuministradorCompleta_BlankDecisionFinal_ReturnsFalse: complete"
    logs(5) = "Test_PCSUB_EsAprobacionSuministradorCompleta_BlankDecisionFinal_ReturnsFalse: PASS"
    Test_PCSUB_EsAprobacionSuministradorCompleta_BlankDecisionFinal_ReturnsFalse = TestHelper.BuildJsonOk(CStr(result), logs)
CleanExit:
    On Error Resume Next
    Call TeardownAll
    Call TestHelper.ResetTestSession
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_EsAprobacionSuministradorCompleta_BlankDecisionFinal_ReturnsFalse: ERR " & Err.Number & " - " & Err.Description
    Test_PCSUB_EsAprobacionSuministradorCompleta_BlankDecisionFinal_ReturnsFalse = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' SLICE E — GuardarDecisionFinal (3 tests)
' ==========================================================================

Public Function Test_PCSUB_GuardarDecisionFinal_PersistsValidDecision() As String
    On Error GoTo EH
    Dim logs() As String: logs = NewLogsArray(6)
    Dim runError As String
    Dim idSol As Long
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim db As DAO.Database

    TestHelper.ForceLocalBackend runError
    If runError <> "" Then
        logs(0) = runError
        Test_PCSUB_GuardarDecisionFinal_PersistsValidDecision = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Call TeardownAll

    idSol = SeedCompletePCSUBFixture()
    Set db = TestHelper.GetTestDb()

    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.decisionFinal = "APROBADO DEFINITIVO"
    pcsub.NombreFirmanteFinal = "Dr. Carlos García"

    svc.GuardarDecisionFinal pcsub, db

    If Not DecisionFieldsMatch(idSol, "APROBADO DEFINITIVO", "Dr. Carlos García") Then
        logs(4) = "Test_PCSUB_GuardarDecisionFinal_PersistsValidDecision: FAILED - decision fields not persisted"
        logs(5) = "Test_PCSUB_GuardarDecisionFinal_PersistsValidDecision: complete"
        Test_PCSUB_GuardarDecisionFinal_PersistsValidDecision = TestHelper.BuildJsonFail("decision fields not persisted", logs)
        GoTo CleanExit
    End If

    logs(0) = "Test_PCSUB_GuardarDecisionFinal_PersistsValidDecision: seed complete fixture E1DEC"
    logs(1) = "Test_PCSUB_GuardarDecisionFinal_PersistsValidDecision: decisionFinal + NombreFirmanteFinal set"
    logs(2) = "Test_PCSUB_GuardarDecisionFinal_PersistsValidDecision: GuardarDecisionFinal succeeded"
    logs(3) = "Test_PCSUB_GuardarDecisionFinal_PersistsValidDecision: DecisionFieldsMatch verified"
    logs(4) = "Test_PCSUB_GuardarDecisionFinal_PersistsValidDecision: complete"
    logs(5) = "Test_PCSUB_GuardarDecisionFinal_PersistsValidDecision: PASS"
    Test_PCSUB_GuardarDecisionFinal_PersistsValidDecision = TestHelper.BuildJsonOk("true", logs)
CleanExit:
    On Error Resume Next
    Call TeardownAll
    Call TestHelper.ResetTestSession
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_GuardarDecisionFinal_PersistsValidDecision: ERR " & Err.Number & " - " & Err.description
    Test_PCSUB_GuardarDecisionFinal_PersistsValidDecision = TestHelper.BuildJsonFail(Err.description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty() As String
    On Error GoTo EH
    Dim logs() As String: logs = NewLogsArray(6)
    Dim runError As String
    Dim idSol As Long
    Dim svc As New DatosPCSUBServicio
    Dim pcsub As DatosPCSUB
    Dim db As DAO.Database
    Dim errCaught As Boolean
    Dim errNumber As Long
    Dim errDescription As String
    Dim expectedDescription As String

    TestHelper.ForceLocalBackend runError
    If runError <> "" Then
        logs(0) = runError
        Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Call TeardownAll

    idSol = SeedCompletePCSUBFixture()
    Set db = TestHelper.GetTestDb()

    Set pcsub = svc.getDatosPCSUB(idSol, db)
    pcsub.decisionFinal = ""  ' Empty!

    On Error Resume Next
    errCaught = False
    svc.GuardarDecisionFinal pcsub, db
    errNumber = Err.Number
    errDescription = Err.description
    If errNumber <> 0 Then errCaught = True
    Err.Clear
    On Error GoTo EH
    expectedDescription = "Error #513: Decisión Final obligatoria." & vbNewLine & "Pila de Llamadas:" & vbNewLine & "   en DatosPCSUBServicio.GuardarDecisionFinal"

    If Not errCaught Or errNumber <> 513 Or errDescription <> expectedDescription Then
        logs(4) = "Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty: FAILED - expected exact error 513"
        logs(5) = "Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty: complete"
        Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty = TestHelper.BuildJsonFail("expected exact error 513 message not raised", logs)
        GoTo CleanExit
    End If

    If Not DecisionFieldsMatch(idSol, "DECISIÓN FINAL", "NOMBRE FIRMANTE") Then
        logs(4) = "Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty: FAILED - validation error changed persisted fields"
        logs(5) = "Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty: complete"
        Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty = TestHelper.BuildJsonFail("validation error changed persisted fields", logs)
        GoTo CleanExit
    End If

    logs(0) = "Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty: seed complete fixture E2EMPTY"
    logs(1) = "Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty: decisionFinal empty"
    logs(2) = "Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty: GuardarDecisionFinal raised exact 513 message as expected"
    logs(3) = "Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty: validation error left persisted decision fields unchanged"
    logs(4) = "Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty: complete"
    logs(5) = "Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty: PASS"
    Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty = TestHelper.BuildJsonOk("true", logs)
CleanExit:
    On Error Resume Next
    Call TeardownAll
    Call TestHelper.ResetTestSession
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty: ERR " & Err.Number & " - " & Err.description
    Test_PCSUB_GuardarDecisionFinal_FailsWhenEmpty = TestHelper.BuildJsonFail(Err.description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse() As String
    On Error GoTo EH
    Dim logs() As String: logs = NewLogsArray(6)
    Dim runError As String
    Dim idSol As Long
    Dim svc As New DatosPCSUBServicio
    Dim result As Boolean
    Dim db As DAO.Database

    TestHelper.ForceLocalBackend runError
    If runError <> "" Then
        logs(0) = runError
        Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Call TeardownAll

    idSol = SeedCompletePCSUBFixture()
    Set db = TestHelper.GetTestDb()

    ' Pre-condition: decisionFinal populated, but NombreFirmanteFinal blank
    db.Execute "UPDATE tbDatosPCSUB SET NombreFirmanteFinal = '' WHERE idSolicitud = " & idSol, dbFailOnError
    If Not DecisionFieldsMatch(idSol, "DECISIÓN FINAL", "") Then
        logs(4) = "Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse: FAILED - precondition not established"
        logs(5) = "Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse: complete"
        Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse = TestHelper.BuildJsonFail("precondition not established", logs)
        GoTo CleanExit
    End If

    result = svc.EsDecisionFinalCompleta(idSol, db)

    If result Then
        logs(4) = "Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse: FAILED - expected fix to return False when NombreFirmanteFinal blank"
        logs(5) = "Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse: complete"
        Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse = TestHelper.BuildJsonFail("expected fix to return False", logs)
        GoTo CleanExit
    End If

    ' FIX verified: EsDecisionFinalCompleta now returns False when NombreFirmanteFinal is blank
    logs(0) = "Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse: seed complete fixture E3GAP"
    logs(1) = "Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse: decisionFinal populated, NombreFirmanteFinal blank"
    logs(2) = "Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse: EsDecisionFinalCompleta returned " & result
    logs(3) = "Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse: FIX verified - returns false when firmante blank"
    logs(4) = "Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse: complete"
    logs(5) = "Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse: PASS"
    Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse = TestHelper.BuildJsonOk(CStr(result), logs)
CleanExit:
    On Error Resume Next
    Call TeardownAll
    Call TestHelper.ResetTestSession
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse: ERR " & Err.Number & " - " & Err.description
    Test_PCSUB_EsDecisionFinalCompleta_MissingFirmante_ReturnsFalse = TestHelper.BuildJsonFail(Err.description, logs)
    Resume CleanExit
End Function

' ==========================================================================
' SLICE F — PurgaTecnica + Eliminar* side-effects (5 tests)
' ==========================================================================

Public Function Test_PCSUB_PurgaTecnica_ClearsAllTecnicaFields() As String
    On Error GoTo EH
    Dim logs() As String: logs = NewLogsArray(7)
    Dim runError As String
    Dim idSol As Long
    Dim svc As New DatosPCSUBServicio
    Dim db As DAO.Database

    TestHelper.ForceLocalBackend runError
    If runError <> "" Then
        logs(0) = runError
        Test_PCSUB_PurgaTecnica_ClearsAllTecnicaFields = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Call TeardownAll

    idSol = SeedPCSUBTechnicalFixture()
    Set db = TestHelper.GetTestDb()

    svc.PurgaTecnica idSol, db

    If Not TecnicaFieldsCleared(idSol) Then
        logs(4) = "Test_PCSUB_PurgaTecnica_ClearsAllTecnicaFields: FAILED - tecnica fields not cleared"
        logs(5) = "Test_PCSUB_PurgaTecnica_ClearsAllTecnicaFields: complete"
        Test_PCSUB_PurgaTecnica_ClearsAllTecnicaFields = TestHelper.BuildJsonFail("tecnica fields not cleared", logs)
        GoTo CleanExit
    End If

    logs(0) = "Test_PCSUB_PurgaTecnica_ClearsAllTecnicaFields: seed complete fixture F1PURGA"
    logs(1) = "Test_PCSUB_PurgaTecnica_ClearsAllTecnicaFields: tecnica fields populated before"
    logs(2) = "Test_PCSUB_PurgaTecnica_ClearsAllTecnicaFields: PurgaTecnica called"
    logs(3) = "Test_PCSUB_PurgaTecnica_ClearsAllTecnicaFields: TecnicaFieldsCleared verified"
    logs(4) = "Test_PCSUB_PurgaTecnica_ClearsAllTecnicaFields: complete"
    logs(5) = "Test_PCSUB_PurgaTecnica_ClearsAllTecnicaFields: PASS"
    logs(6) = "Test_PCSUB_PurgaTecnica_ClearsAllTecnicaFields: descripcionMaterialAfectado + propuesta + impacto cleared"
    Test_PCSUB_PurgaTecnica_ClearsAllTecnicaFields = TestHelper.BuildJsonOk("true", logs)
CleanExit:
    On Error Resume Next
    Call TeardownAll
    Call TestHelper.ResetTestSession
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_PurgaTecnica_ClearsAllTecnicaFields: ERR " & Err.Number & " - " & Err.description
    Test_PCSUB_PurgaTecnica_ClearsAllTecnicaFields = TestHelper.BuildJsonFail(Err.description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields() As String
    On Error GoTo EH
    Dim logs() As String: logs = NewLogsArray(8)
    Dim runError As String
    Dim idSol As Long
    Dim svc As New DatosPCSUBServicio
    Dim db As DAO.Database

    TestHelper.ForceLocalBackend runError
    If runError <> "" Then
        logs(0) = runError
        Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Call TeardownAll

    idSol = SeedCompletePCSUBFixture()
    Set db = TestHelper.GetTestDb()

    ' Pre-condition: verify all groups are populated
    If Not RacFieldsMatch(idSol, "RAC-2024-001", "APROBADO") Then
        logs(0) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: FAILED - RAC not seeded"
        logs(1) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: complete"
        Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields = TestHelper.BuildJsonFail("RAC not seeded", logs)
        GoTo CleanExit
    End If
    If Not AprobacionFieldsMatch(idSol, "FIRMA OFICINA TECNICA", "FIRMA REP SUB") Then
        logs(0) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: FAILED - aprobacion not seeded"
        logs(1) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: complete"
        Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields = TestHelper.BuildJsonFail("aprobacion not seeded", logs)
        GoTo CleanExit
    End If

    svc.EliminarDictamenRAC idSol, db

    ' Verify RAC cleared
    If Not RacFieldsCleared(idSol) Then
        logs(4) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: FAILED - RAC fields not cleared"
        logs(5) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: complete"
        Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields = TestHelper.BuildJsonFail("RAC fields not cleared", logs)
        GoTo CleanExit
    End If

    ' Verify Aprobacion unchanged (side-effect check)
    If Not AprobacionFieldsMatch(idSol, "FIRMA OFICINA TECNICA", "FIRMA REP SUB") Then
        logs(4) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: FAILED - aprobacion fields were mutated"
        logs(5) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: complete"
        Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields = TestHelper.BuildJsonFail("aprobacion fields mutated", logs)
        GoTo CleanExit
    End If

    logs(0) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: seed complete fixture F2RAC"
    logs(1) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: RAC + Aprobacion fields populated"
    logs(2) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: EliminarDictamenRAC called"
    logs(3) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: RacFieldsCleared verified"
    logs(4) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: AprobacionFieldsMatch verified (side-effect check)"
    logs(5) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: complete"
    logs(6) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: PASS"
    logs(7) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: only RAC fields cleared, rest unchanged"
    Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields = TestHelper.BuildJsonOk("true", logs)
CleanExit:
    On Error Resume Next
    Call TeardownAll
    Call TestHelper.ResetTestSession
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields: ERR " & Err.Number & " - " & Err.description
    Test_PCSUB_EliminarDictamenRAC_ClearsOnlyRACFields = TestHelper.BuildJsonFail(Err.description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields() As String
    On Error GoTo EH
    Dim logs() As String: logs = NewLogsArray(8)
    Dim runError As String
    Dim idSol As Long
    Dim svc As New DatosPCSUBServicio
    Dim db As DAO.Database

    TestHelper.ForceLocalBackend runError
    If runError <> "" Then
        logs(0) = runError
        Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Call TeardownAll

    idSol = SeedCompletePCSUBFixture()
    Set db = TestHelper.GetTestDb()

    ' Pre-condition: verify all groups are populated
    If Not AprobacionFieldsMatch(idSol, "FIRMA OFICINA TECNICA", "FIRMA REP SUB") Then
        logs(0) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: FAILED - aprobacion not seeded"
        logs(1) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: complete"
        Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields = TestHelper.BuildJsonFail("aprobacion not seeded", logs)
        GoTo CleanExit
    End If
    If Not RacFieldsMatch(idSol, "RAC-2024-001", "APROBADO") Then
        logs(0) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: FAILED - RAC not seeded"
        logs(1) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: complete"
        Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields = TestHelper.BuildJsonFail("RAC not seeded", logs)
        GoTo CleanExit
    End If

    svc.EliminarAprobacionSuministrador idSol, db

    ' Verify Aprobacion cleared
    If Not AprobacionFieldsCleared(idSol) Then
        logs(4) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: FAILED - signer fields not cleared"
        logs(5) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: complete"
        Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields = TestHelper.BuildJsonFail("signer fields not cleared", logs)
        GoTo CleanExit
    End If

    ' Verify RAC unchanged (side-effect check)
    If Not RacFieldsMatch(idSol, "RAC-2024-001", "APROBADO") Then
        logs(4) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: FAILED - RAC fields were mutated"
        logs(5) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: complete"
        Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields = TestHelper.BuildJsonFail("RAC fields mutated", logs)
        GoTo CleanExit
    End If

    logs(0) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: seed complete fixture F3APROB"
    logs(1) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: Aprobacion + RAC fields populated"
    logs(2) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: EliminarAprobacionSuministrador called"
    logs(3) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: AprobacionFieldsCleared verified"
    logs(4) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: RacFieldsMatch verified (side-effect check)"
    logs(5) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: complete"
    logs(6) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: PASS"
    logs(7) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: only signer fields cleared, rest unchanged"
    Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields = TestHelper.BuildJsonOk("true", logs)
CleanExit:
    On Error Resume Next
    Call TeardownAll
    Call TestHelper.ResetTestSession
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields: ERR " & Err.Number & " - " & Err.description
    Test_PCSUB_EliminarAprobacionSuministrador_ClearsOnlySignerFields = TestHelper.BuildJsonFail(Err.description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields() As String
    On Error GoTo EH
    Dim logs() As String: logs = NewLogsArray(8)
    Dim runError As String
    Dim idSol As Long
    Dim svc As New DatosPCSUBServicio
    Dim db As DAO.Database

    TestHelper.ForceLocalBackend runError
    If runError <> "" Then
        logs(0) = runError
        Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Call TeardownAll

    idSol = SeedCompletePCSUBFixture()
    Set db = TestHelper.GetTestDb()

    ' Pre-condition: verify all groups are populated
    If Not DecisionFieldsMatch(idSol, "DECISIÓN FINAL", "NOMBRE FIRMANTE") Then
        logs(0) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: FAILED - decision not seeded"
        logs(1) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: complete"
        Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields = TestHelper.BuildJsonFail("decision not seeded", logs)
        GoTo CleanExit
    End If
    If Not RacFieldsMatch(idSol, "RAC-2024-001", "APROBADO") Then
        logs(0) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: FAILED - RAC not seeded"
        logs(1) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: complete"
        Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields = TestHelper.BuildJsonFail("RAC not seeded", logs)
        GoTo CleanExit
    End If

    svc.EliminarDecisionFinal idSol, db

    ' Verify Decision cleared
    If Not DecisionFieldsCleared(idSol) Then
        logs(4) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: FAILED - decision fields not cleared"
        logs(5) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: complete"
        Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields = TestHelper.BuildJsonFail("decision fields not cleared", logs)
        GoTo CleanExit
    End If

    ' Verify RAC unchanged (side-effect check)
    If Not RacFieldsMatch(idSol, "RAC-2024-001", "APROBADO") Then
        logs(4) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: FAILED - RAC fields were mutated"
        logs(5) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: complete"
        Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields = TestHelper.BuildJsonFail("RAC fields mutated", logs)
        GoTo CleanExit
    End If

    logs(0) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: seed complete fixture F4DEC"
    logs(1) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: Decision + RAC fields populated"
    logs(2) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: EliminarDecisionFinal called"
    logs(3) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: DecisionFieldsCleared verified"
    logs(4) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: RacFieldsMatch verified (side-effect check)"
    logs(5) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: complete"
    logs(6) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: PASS"
    logs(7) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: only decision fields cleared, rest unchanged"
    Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields = TestHelper.BuildJsonOk("true", logs)
CleanExit:
    On Error Resume Next
    Call TeardownAll
    Call TestHelper.ResetTestSession
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields: ERR " & Err.Number & " - " & Err.description
    Test_PCSUB_EliminarDecisionFinal_ClearsOnlyDecisionFields = TestHelper.BuildJsonFail(Err.description, logs)
    Resume CleanExit
End Function

Public Function Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow() As String
    On Error GoTo EH
    Dim logs() As String: logs = NewLogsArray(7)
    Dim runError As String
    Dim idSol As Long
    Dim svc As New DatosPCSUBServicio
    Dim db As DAO.Database
    Dim countBefore As Long
    Dim countAfter As Long

    TestHelper.ForceLocalBackend runError
    If runError <> "" Then
        logs(0) = runError
        Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Call TeardownAll

    idSol = SeedCompletePCSUBFixture()
    Set db = TestHelper.GetTestDb()

    countBefore = CountPCSUBRows(idSol)
    If countBefore < 1 Then
        logs(0) = "Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow: FAILED - no PCSUB row to delete"
        logs(1) = "Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow: complete"
        Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow = TestHelper.BuildJsonFail("no PCSUB row to delete", logs)
        GoTo CleanExit
    End If

    DatosPCSUBRepositorio.EliminarPorIdSolicitud idSol, db

    countAfter = CountPCSUBRows(idSol)

    If countAfter <> 0 Then
        logs(4) = "Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow: FAILED - row not deleted"
        logs(5) = "Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow: complete"
        Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow = TestHelper.BuildJsonFail("row not deleted", logs)
        GoTo CleanExit
    End If

    logs(0) = "Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow: seed complete fixture F5DEL"
    logs(1) = "Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow: countBefore=" & countBefore
    logs(2) = "Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow: EliminarPorIdSolicitud called"
    logs(3) = "Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow: countAfter=" & countAfter & " (expected 0)"
    logs(4) = "Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow: row deleted confirmed"
    logs(5) = "Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow: complete"
    logs(6) = "Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow: PASS"
    Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow = TestHelper.BuildJsonOk("true", logs)
CleanExit:
    On Error Resume Next
    Call TeardownAll
    Call TestHelper.ResetTestSession
    Exit Function

EH:
    logs(UBound(logs)) = "Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow: ERR " & Err.Number & " - " & Err.description
    Test_PCSUB_EliminarPorIdSolicitud_DeletesPCSUBRow = TestHelper.BuildJsonFail(Err.description, logs)
    Resume CleanExit
End Function




