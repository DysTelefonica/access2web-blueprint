Attribute VB_Name = "Test_CalidadCacheRiesgo"
Option Compare Database
Option Explicit

Private Const FIX_EXPEDIENTE_ID As Long = 901500
Private Const FIX_PROYECTO_ID As Long = 901501
Private Const FIX_EDICION_ID As Long = 901502
Private Const FIX_RIESGO_ID As Long = 901503
Private Const FIX_RIESGO_UNRELATED_ID As Long = 901504
Private Const FIX_CODIGO_RIESGO As String = "T9001"
Private Const FIX_CODIGO_RIESGO_UNRELATED As String = "T9002"
Private Const FIX_RETIPIFICATION_CODE As String = "R-PR-30"
Private Const FIX_RETIPIFICATION_DATE_SQL As String = "#01/15/2026#"
Private Const FIX_DETECTED_DATE_SQL As String = "#01/10/2026#"
Private Const FIX_TASK_USER_NAME As String = "TEST_CACHE_CALIDAD"
Private Const FIX_AUTHORIZED_NAME As String = "TEST_CACHE_AUTORIZADO"
Private Const FIX_RISK_ACTOR As String = "TEST_CACHE_ACTOR"
Private Const FIX_UPDATED_DESCRIPTION As String = "Fixture riesgo actualizado determinista"

' Schema-first fixture graph verified with Dysflow/ERD:
' TbExpedientes(IDExpediente) -> TbProyectos(IDProyecto, IDExpediente, NombreUsuarioCalidad, CadenaNombreAutorizados)
' -> TbProyectosEdiciones(IDEdicion, IDProyecto, Edicion) -> TbRiesgos(IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, retipification fields).
' Teardown deletes child tables/FKs first, then risks, editions, projects, and expediente.

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

Private Function SqlText(ByVal value As String) As String
    SqlText = "'" & Replace(value, "'", "''") & "'"
End Function

Private Function FixtureRiskKey() As String
    FixtureRiskKey = CStr(FIX_RIESGO_ID)
End Function

Private Function FixtureUnrelatedRiskKey() As String
    FixtureUnrelatedRiskKey = CStr(FIX_RIESGO_UNRELATED_ID)
End Function

Private Function CountFixtureOwnedPendingKeys(ByVal pending As Scripting.Dictionary) As Long
    If pending Is Nothing Then Exit Function

    If pending.Exists(FixtureRiskKey()) Then
        CountFixtureOwnedPendingKeys = CountFixtureOwnedPendingKeys + 1
    End If
    If pending.Exists(FixtureUnrelatedRiskKey()) Then
        CountFixtureOwnedPendingKeys = CountFixtureOwnedPendingKeys + 1
    End If
End Function

Private Function ValidatePendingFixtureKeys( _
    ByVal pending As Scripting.Dictionary, _
    ByVal expectedCount As Long, _
    Optional ByRef p_Error As String _
) As Boolean
    p_Error = ""
    If pending Is Nothing Then
        p_Error = "ColRiesgosPorReTipificar no puede ser Nothing"
        Exit Function
    End If

    If CountFixtureOwnedPendingKeys(pending) <> expectedCount Then
        p_Error = "Claves fixture pendientes esperado=" & CStr(expectedCount) & _
            " real=" & CStr(CountFixtureOwnedPendingKeys(pending))
        Exit Function
    End If
    If expectedCount = 1 And Not pending.Exists(FixtureRiskKey()) Then
        p_Error = "La cache de tareas debe contener el riesgo fixture determinista"
        Exit Function
    End If
    If expectedCount = 0 And pending.Exists(FixtureRiskKey()) Then
        p_Error = "La cache de tareas no debe contener el riesgo fixture determinista"
        Exit Function
    End If
    If pending.Exists(FixtureUnrelatedRiskKey()) Then
        p_Error = "La cache de tareas no debe contener el riesgo fixture no relacionado"
        Exit Function
    End If

    ValidatePendingFixtureKeys = True
End Function

Private Function ResultAfterTeardown( _
    ByVal resultJson As String, _
    ByVal teardownError As String, _
    ByRef logs() As String _
) As String
    ResultAfterTeardown = resultJson
    If teardownError = "" Then Exit Function

    If resultJson = "" Or InStr(1, resultJson, """ok"":true", vbTextCompare) > 0 Then
        ResultAfterTeardown = BuildFail(teardownError, logs)
    End If
End Function

Private Function FixtureDb(Optional ByRef p_Error As String) As DAO.Database
    Set FixtureDb = Test_Fixtures.GetTestDb(p_Error)
    If FixtureDb Is Nothing And p_Error = "" Then
        p_Error = "GetTestDb devolvió Nothing"
    End If
End Function

Private Function TeardownCalidadCacheGraph(Optional ByRef p_Error As String) As Boolean
    On Error GoTo EH

    Dim db As DAO.Database
    Dim dbErr As String
    Dim cacheErr As String
    Dim ids As String

    p_Error = ""
    ids = CStr(FIX_RIESGO_ID) & "," & CStr(FIX_RIESGO_UNRELATED_ID)
    Set db = FixtureDb(dbErr)
    If db Is Nothing Then
        If dbErr = "" Then dbErr = "GetTestDb devolvió Nothing"
        Err.Raise 1000, "TeardownCalidadCacheGraph", dbErr
    End If

    db.Execute "DELETE FROM TbTareas WHERE IDProyecto=" & FIX_PROYECTO_ID & " OR IDRiesgo IN (" & ids & ")", dbFailOnError
    db.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE IDProyecto=" & FIX_PROYECTO_ID & " OR IDEdicion=" & FIX_EDICION_ID, dbFailOnError
    db.Execute "DELETE FROM TbRiesgosPlanContingenciaDetalle WHERE IDContingencia IN " & _
        "(SELECT IDContingencia FROM TbRiesgosPlanContingenciaPpal WHERE IDRiesgo IN (" & ids & "))", dbFailOnError
    db.Execute "DELETE FROM TbRiesgosPlanMitigacionDetalle WHERE IDMitigacion IN " & _
        "(SELECT IDMitigacion FROM TbRiesgosPlanMitigacionPpal WHERE IDRiesgo IN (" & ids & "))", dbFailOnError
    db.Execute "DELETE FROM TbRiesgosPlanContingenciaPpal WHERE IDRiesgo IN (" & ids & ")", dbFailOnError
    db.Execute "DELETE FROM TbRiesgosPlanMitigacionPpal WHERE IDRiesgo IN (" & ids & ")", dbFailOnError
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo IN (" & ids & ")", dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_ID, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_PROYECTO_ID, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_EXPEDIENTE_ID, dbFailOnError
    TeardownCalidadCacheGraph = True

Cleanup:
    On Error Resume Next
    cacheErr = ""
    InvalidarCacheRiesgo CStr(FIX_RIESGO_ID), cacheErr
    If cacheErr <> "" And p_Error = "" Then p_Error = "InvalidarCacheRiesgo fixture: " & cacheErr
    cacheErr = ""
    InvalidarCacheRiesgo CStr(FIX_RIESGO_UNRELATED_ID), cacheErr
    If cacheErr <> "" And p_Error = "" Then p_Error = "InvalidarCacheRiesgo unrelated: " & cacheErr
    cacheErr = ""
    InvalidarCacheEdicion CStr(FIX_EDICION_ID), cacheErr
    If cacheErr <> "" And p_Error = "" Then p_Error = "InvalidarCacheEdicion fixture: " & cacheErr
    If p_Error <> "" Then TeardownCalidadCacheGraph = False
    Set m_ObjTareasCalidad = Nothing
    Set db = Nothing
    On Error GoTo 0
    Exit Function
EH:
    p_Error = "TeardownCalidadCacheGraph: " & Err.Description
    Resume Cleanup
End Function

Private Sub InsertRiskFixture( _
    ByVal db As DAO.Database, _
    ByVal riskId As Long, _
    ByVal riskCode As String, _
    ByVal libraryCode As String, _
    ByVal pendingRetipification As Boolean, _
    ByVal description As String _
)
    Dim requiresLibrary As String
    Dim pendingText As String
    Dim retipDateSql As String
    Dim sql As String

    If pendingRetipification Then
        requiresLibrary = "Sí"
        pendingText = "Sí"
        retipDateSql = FIX_RETIPIFICATION_DATE_SQL
    Else
        requiresLibrary = "No"
        pendingText = "No"
        retipDateSql = "Null"
    End If

    sql = "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, DetectadoPor, EntidadDetecta, Plazo, Calidad, Coste, " & _
        "ImpactoGlobal, Vulnerabilidad, Valoracion, Mitigacion, Contingencia, " & _
        "RequierePlanContingencia, Descripcion, CausaRaiz, Estado, Priorizacion, " & _
        "RequiereRiesgoDeBiblioteca, CodRiesgoBiblioteca, RiesgoPendienteRetipificacion, " & _
        "FechaRiesgoParaRetipificar) VALUES (" & riskId & ", " & FIX_EDICION_ID & ", " & _
        SqlText("CACHE-FIX-" & CStr(riskId)) & ", " & SqlText(riskCode) & ", " & FIX_DETECTED_DATE_SQL & ", " & _
        SqlText(FIX_RISK_ACTOR) & ", " & SqlText("TEST") & ", " & SqlText("Medio") & ", " & _
        SqlText("Medio") & ", " & SqlText("Medio") & ", " & SqlText("Medio") & ", " & _
        SqlText("Medio") & ", " & SqlText("Medio") & ", " & SqlText("Reducir") & ", " & _
        SqlText("No") & ", " & SqlText("No") & ", " & SqlText(description) & ", " & _
        SqlText("Fixture") & ", " & SqlText("Detectado") & ", 3, " & SqlText(requiresLibrary) & ", " & _
        SqlText(libraryCode) & ", " & SqlText(pendingText) & ", " & retipDateSql & ")"
    db.Execute sql, dbFailOnError
End Sub

Private Function FixturePendingRetipificationCount(Optional ByRef p_Error As String) As Long
    On Error GoTo EH

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    p_Error = ""
    Set db = FixtureDb(p_Error)
    If db Is Nothing Then GoTo Cleanup

    sql = "SELECT Count(*) AS FixtureCount " & _
        "FROM (TbProyectos INNER JOIN TbProyectosEdiciones " & _
        "ON TbProyectos.IDProyecto = TbProyectosEdiciones.IDProyecto) " & _
        "INNER JOIN TbRiesgos ON TbProyectosEdiciones.IDEdicion = TbRiesgos.IDEdicion " & _
        "WHERE TbProyectos.IDProyecto=" & FIX_PROYECTO_ID & " " & _
        "AND TbProyectosEdiciones.IDEdicion=" & FIX_EDICION_ID & " " & _
        "AND TbRiesgos.IDRiesgo=" & FIX_RIESGO_ID & " " & _
        "AND TbRiesgos.CodRiesgoBiblioteca=" & SqlText(FIX_RETIPIFICATION_CODE) & " " & _
        "AND TbProyectosEdiciones.FechaPublicacion Is Null " & _
        "AND TbProyectos.NombreUsuarioCalidad=" & SqlText(FIX_TASK_USER_NAME) & " " & _
        "AND TbProyectos.CadenaNombreAutorizados Like " & SqlText("*" & FIX_AUTHORIZED_NAME & "*")
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then FixturePendingRetipificationCount = CLng(rs.Fields("FixtureCount").Value)

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0
    Exit Function
EH:
    p_Error = "FixturePendingRetipificationCount: " & Err.Description
    Resume Cleanup
End Function

Private Function ValidateFixturePendingRetipification( _
    ByVal expectedCount As Long, _
    Optional ByRef p_Error As String _
) As Boolean
    Dim actualCount As Long

    p_Error = ""
    actualCount = FixturePendingRetipificationCount(p_Error)
    If p_Error <> "" Then Exit Function
    If actualCount <> expectedCount Then
        p_Error = "Fixture retipificación esperado=" & CStr(expectedCount) & " real=" & CStr(actualCount)
        Exit Function
    End If
    ValidateFixturePendingRetipification = True
End Function

Private Function SeedCalidadCacheGraph( _
    ByVal pendingRetipification As Boolean, _
    ByRef expectedFixturePendingCount As Long, _
    Optional ByRef p_Error As String _
) As Boolean
    On Error GoTo EH

    Dim db As DAO.Database
    Dim libraryCode As String
    Dim sql As String

    p_Error = ""
    expectedFixturePendingCount = 0

    If Not TeardownCalidadCacheGraph(p_Error) Then GoTo Cleanup
    Set m_ObjTareasCalidad = Nothing
    If Not ValidateFixturePendingRetipification(0, p_Error) Then GoTo Cleanup

    If pendingRetipification Then
        libraryCode = FIX_RETIPIFICATION_CODE
    Else
        libraryCode = "CACHE-FIXTURE-NORMAL"
    End If

    Set db = FixtureDb(p_Error)
    If db Is Nothing Then GoTo Cleanup

    sql = "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
        "VALUES (" & FIX_EXPEDIENTE_ID & ", " & SqlText("TSTCACHE") & ", " & _
        SqlText("Fixture cache calidad") & ", " & SqlText("Test") & ", 1)"
    db.Execute sql, dbFailOnError

    sql = "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto, NombreProyecto, " & _
        "NombreUsuarioCalidad, CadenaNombreAutorizados, ParaInformeAvisos) VALUES (" & _
        FIX_PROYECTO_ID & ", " & FIX_EXPEDIENTE_ID & ", " & SqlText("TST-CACHE") & ", " & _
        SqlText("Fixture cache calidad") & ", " & SqlText(FIX_TASK_USER_NAME) & ", " & _
        SqlText(FIX_AUTHORIZED_NAME) & ", " & SqlText("No") & ")"
    db.Execute sql, dbFailOnError

    sql = "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
        "VALUES (" & FIX_EDICION_ID & ", " & FIX_PROYECTO_ID & ", 1, " & SqlText(FIX_RISK_ACTOR) & ")"
    db.Execute sql, dbFailOnError

    InsertRiskFixture db, FIX_RIESGO_ID, FIX_CODIGO_RIESGO, libraryCode, _
        pendingRetipification, "Fixture riesgo cache calidad"
    InsertRiskFixture db, FIX_RIESGO_UNRELATED_ID, FIX_CODIGO_RIESGO_UNRELATED, _
        "CACHE-FIXTURE-UNRELATED", False, "Fixture riesgo no relacionado"

    expectedFixturePendingCount = IIf(pendingRetipification, 1, 0)
    If Not ValidateFixturePendingRetipification(expectedFixturePendingCount, p_Error) Then GoTo Cleanup
    Set m_ObjTareasCalidad = Nothing
    SeedCalidadCacheGraph = True

Cleanup:
    Set db = Nothing
    Exit Function
EH:
    p_Error = "SeedCalidadCacheGraph: " & Err.Description
    Resume Cleanup
End Function

Private Function RiskDescriptionFromDb(ByVal riskId As Long, Optional ByRef p_Error As String) As String
    On Error GoTo EH

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    p_Error = ""
    Set db = FixtureDb(p_Error)
    If db Is Nothing Then GoTo Cleanup

    sql = "SELECT Descripcion FROM TbRiesgos WHERE IDRiesgo=" & riskId
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Then
        p_Error = "No existe el riesgo fixture " & CStr(riskId)
        GoTo Cleanup
    End If
    RiskDescriptionFromDb = Nz(rs.Fields("Descripcion").Value, "")

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0
    Exit Function
EH:
    p_Error = "RiskDescriptionFromDb: " & Err.Description
    Resume Cleanup
End Function

Public Function Test_CalidadCacheRiesgo_SaveAndRefresh_Contract() As String
    On Error GoTo EH

    Dim logs(0 To 8) As String
    Dim errMsg As String
    Dim teardownError As String
    Dim service As CalidadPublicacionService
    Dim result As CalidadPublicacionResultado
    Dim expectedFixturePendingCount As Long
    Dim pendingAfter As Scripting.Dictionary

    logs(0) = "1. Arrange: ForceLocalBackend uses BackendSandbox"
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_Contract = BuildFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Arrange: seed deterministic pending retipification graph"
    If Not SeedCalidadCacheGraph(True, expectedFixturePendingCount, errMsg) Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_Contract = BuildFail(errMsg, logs)
        GoTo Cleanup
    End If

    logs(2) = "3. Act: service SaveAndRefreshTareasCalidad"
    Set service = New CalidadPublicacionService
    Set result = service.SaveAndRefreshTareasCalidad( _
        CStr(FIX_EDICION_ID), _
        CStr(FIX_RIESGO_ID), _
        "calidad", _
        errMsg)

    logs(3) = "4. Assert: p_Error remains empty and result exists"
    If errMsg <> "" Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_Contract = BuildFail(errMsg, logs)
        GoTo Cleanup
    End If
    If result Is Nothing Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_Contract = BuildFail("SaveAndRefreshTareasCalidad devolvio Nothing", logs)
        GoTo Cleanup
    End If

    logs(4) = "5. Assert: cache invalidation and edition rehydrate flags"
    If result.RiesgoInvalidado <> EnumSiNo.Sí Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_Contract = BuildFail("RiesgoInvalidado debe ser Sí", logs)
        GoTo Cleanup
    End If
    If result.EdicionRehidratada <> EnumSiNo.Sí Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_Contract = BuildFail("EdicionRehidratada debe ser Sí", logs)
        GoTo Cleanup
    End If

    logs(5) = "6. Assert: fixture-owned DB pending count remains exact"
    If Not ValidateFixturePendingRetipification(expectedFixturePendingCount, errMsg) Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_Contract = BuildFail(errMsg, logs)
        GoTo Cleanup
    End If

    logs(6) = "7. Assert: quality task cache exists after refresh"
    If m_ObjTareasCalidad Is Nothing Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_Contract = BuildFail("m_ObjTareasCalidad debe quedar rehidratado", logs)
        GoTo Cleanup
    End If
    Set pendingAfter = m_ObjTareasCalidad.ColRiesgosPorReTipificar
    If pendingAfter Is Nothing Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_Contract = BuildFail("ColRiesgosPorReTipificar no puede ser Nothing", logs)
        GoTo Cleanup
    End If

    logs(7) = "8. Assert: fixture-owned pending cache keys are exact"
    If Not ValidatePendingFixtureKeys(pendingAfter, expectedFixturePendingCount, errMsg) Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_Contract = BuildFail(errMsg, logs)
        GoTo Cleanup
    End If

    logs(8) = "9. Assert: deterministic fixture graph drove the contract"
    Test_CalidadCacheRiesgo_SaveAndRefresh_Contract = BuildOk("save_and_refresh_contract", logs)

Cleanup:
    On Error Resume Next
    Set pendingAfter = Nothing
    Set result = Nothing
    Set service = Nothing
    teardownError = ""
    If Not TeardownCalidadCacheGraph(teardownError) Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_Contract = ResultAfterTeardown(Test_CalidadCacheRiesgo_SaveAndRefresh_Contract, teardownError, logs)
    End If
    Test_Helper.ResetTestSession errMsg
    Exit Function
EH:
    Test_CalidadCacheRiesgo_SaveAndRefresh_Contract = BuildFail("Test_CalidadCacheRiesgo_SaveAndRefresh_Contract: " & Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_CalidadCacheRiesgo_SaveAndRefresh_PropagatesError() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    Dim errMsg As String
    Dim pendingBefore As Scripting.Dictionary
    Dim service As CalidadPublicacionService
    Dim result As CalidadPublicacionResultado

    logs(0) = "1. Arrange: ForceLocalBackend uses BackendSandbox"
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_PropagatesError = BuildFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Arrange: existing pending tasks collection"
    Set pendingBefore = New Scripting.Dictionary
    pendingBefore.Add "ERR-RISK", "kept-before-error"
    Set m_ObjTareasCalidad = New TareasCalidad
    Set m_ObjTareasCalidad.ColRiesgosPorReTipificar = pendingBefore

    logs(2) = "3. Act: service receives invalid risk/edition ids"
    Set service = New CalidadPublicacionService
    Set result = service.SaveAndRefreshTareasCalidad("0", "0", "calidad", errMsg)

    logs(3) = "4. Assert: p_Error is propagated and result is safe"
    If errMsg = "" Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_PropagatesError = BuildFail("SaveAndRefreshTareasCalidad debe propagar p_Error", logs)
        GoTo Cleanup
    End If
    If result Is Nothing Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_PropagatesError = BuildFail("El resultado de error debe ser un objeto seguro", logs)
        GoTo Cleanup
    End If

    logs(4) = "5. Assert: previous pending task state is preserved exactly"
    If m_ObjTareasCalidad Is Nothing Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_PropagatesError = BuildFail("m_ObjTareasCalidad no debe limpiarse cuando hay error", logs)
        GoTo Cleanup
    End If
    If m_ObjTareasCalidad.ColRiesgosPorReTipificar.Count <> 1 Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_PropagatesError = BuildFail("La colección previa debe mantener cardinalidad 1", logs)
        GoTo Cleanup
    End If
    If m_ObjTareasCalidad.ColRiesgosPorReTipificar("ERR-RISK") <> "kept-before-error" Then
        Test_CalidadCacheRiesgo_SaveAndRefresh_PropagatesError = BuildFail("La colección previa no debe mutar cuando hay error", logs)
        GoTo Cleanup
    End If

    Test_CalidadCacheRiesgo_SaveAndRefresh_PropagatesError = BuildOk("save_and_refresh_error_propagated", logs)

Cleanup:
    On Error Resume Next
    Set m_ObjTareasCalidad = Nothing
    Set pendingBefore = Nothing
    Set result = Nothing
    Set service = Nothing
    Test_Helper.ResetTestSession errMsg
    Exit Function
EH:
    Test_CalidadCacheRiesgo_SaveAndRefresh_PropagatesError = BuildFail("Test_CalidadCacheRiesgo_SaveAndRefresh_PropagatesError: " & Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_CalidadCacheRiesgo_InvalidarScopeFull_RehidrataEdicion() As String
    On Error GoTo EH

    Dim logs(0 To 8) As String
    Dim errMsg As String
    Dim teardownError As String
    Dim edicionAntes As Edicion
    Dim edicionDespues As Edicion
    Dim riesgoAntes As riesgo
    Dim riesgoDespues As riesgo
    Dim riesgoUnrelatedAntes As riesgo
    Dim riesgoUnrelatedDespues As riesgo
    Dim expectedFixturePendingCount As Long

    logs(0) = "1. Arrange: ForceLocalBackend uses BackendSandbox"
    ResetGlobals errMsg
    If errMsg <> "" Then
        Test_CalidadCacheRiesgo_InvalidarScopeFull_RehidrataEdicion = BuildFail("ResetGlobals before arrange failed: " & errMsg, logs)
        Exit Function
    End If

    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_CalidadCacheRiesgo_InvalidarScopeFull_RehidrataEdicion = BuildFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Arrange: seed deterministic target and unrelated risks"
    If Not SeedCalidadCacheGraph(False, expectedFixturePendingCount, errMsg) Then
        Test_CalidadCacheRiesgo_InvalidarScopeFull_RehidrataEdicion = BuildFail(errMsg, logs)
        GoTo Cleanup
    End If

    logs(2) = "3. Arrange: cache edition, target risk, and unrelated risk"
    Set edicionAntes = GetCachedEdicion(CStr(FIX_EDICION_ID), errMsg)
    If errMsg <> "" Then GoTo FailWithErr
    Set riesgoAntes = GetCachedRiesgo(CStr(FIX_RIESGO_ID), errMsg)
    If errMsg <> "" Then GoTo FailWithErr
    Set riesgoUnrelatedAntes = GetCachedRiesgo(CStr(FIX_RIESGO_UNRELATED_ID), errMsg)
    If errMsg <> "" Then GoTo FailWithErr
    If edicionAntes Is Nothing Or riesgoAntes Is Nothing Or riesgoUnrelatedAntes Is Nothing Then
        Test_CalidadCacheRiesgo_InvalidarScopeFull_RehidrataEdicion = BuildFail("No se pudo preparar cache deterministic", logs)
        GoTo Cleanup
    End If

    logs(3) = "4. Act: helper invalidates full scope and rehydrates edition"
    Set edicionDespues = Application.Run( _
        "InvalidarYRehidratarRiesgoPorScope", _
        CStr(FIX_EDICION_ID), _
        CStr(FIX_RIESGO_ID), _
        "full", _
        errMsg)

    logs(4) = "5. Assert: helper propagates no error and returns edition"
    If errMsg <> "" Then GoTo FailWithErr
    If edicionDespues Is Nothing Then
        Test_CalidadCacheRiesgo_InvalidarScopeFull_RehidrataEdicion = BuildFail("La edición rehidratada no puede ser Nothing", logs)
        GoTo Cleanup
    End If

    logs(5) = "6. Assert: returned edition is the requested fixture edition"
    If CStr(edicionDespues.IDEdicion) <> CStr(FIX_EDICION_ID) Then
        Test_CalidadCacheRiesgo_InvalidarScopeFull_RehidrataEdicion = BuildFail("La edición rehidratada no corresponde al fixture", logs)
        GoTo Cleanup
    End If

    logs(6) = "7. Assert: full scope replaces stale cached edition reference"
    If edicionDespues Is edicionAntes Then
        Test_CalidadCacheRiesgo_InvalidarScopeFull_RehidrataEdicion = BuildFail("Full scope debe invalidar y rehidratar una nueva instancia de edición", logs)
        GoTo Cleanup
    End If

    logs(7) = "8. Assert: target risk cache is replaced"
    Set riesgoDespues = GetCachedRiesgo(CStr(FIX_RIESGO_ID), errMsg)
    If errMsg <> "" Then GoTo FailWithErr
    If riesgoDespues Is Nothing Or riesgoDespues Is riesgoAntes Then
        Test_CalidadCacheRiesgo_InvalidarScopeFull_RehidrataEdicion = BuildFail("Full scope debe invalidar el riesgo objetivo", logs)
        GoTo Cleanup
    End If

    logs(8) = "9. Assert: unrelated cached risk is preserved"
    Set riesgoUnrelatedDespues = GetCachedRiesgo(CStr(FIX_RIESGO_UNRELATED_ID), errMsg)
    If errMsg <> "" Then GoTo FailWithErr
    If Not (riesgoUnrelatedDespues Is riesgoUnrelatedAntes) Then
        Test_CalidadCacheRiesgo_InvalidarScopeFull_RehidrataEdicion = BuildFail("Full scope no debe eliminar la cache de otro riesgo", logs)
        GoTo Cleanup
    End If

    Test_CalidadCacheRiesgo_InvalidarScopeFull_RehidrataEdicion = BuildOk("full_scope_rehydrates_edition", logs)
    GoTo Cleanup

FailWithErr:
    Test_CalidadCacheRiesgo_InvalidarScopeFull_RehidrataEdicion = BuildFail(errMsg, logs)

Cleanup:
    On Error Resume Next
    teardownError = ""
    If Not TeardownCalidadCacheGraph(teardownError) Then
        Test_CalidadCacheRiesgo_InvalidarScopeFull_RehidrataEdicion = ResultAfterTeardown(Test_CalidadCacheRiesgo_InvalidarScopeFull_RehidrataEdicion, teardownError, logs)
    End If
    errMsg = ""
    ResetGlobals errMsg
    Test_Helper.ResetTestSession errMsg
    Exit Function
EH:
    Test_CalidadCacheRiesgo_InvalidarScopeFull_RehidrataEdicion = BuildFail("Test_CalidadCacheRiesgo_InvalidarScopeFull_RehidrataEdicion: " & Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_CalidadCacheRiesgo_TareasRefresh_LimpiaDiccionarios() As String
    On Error GoTo EH

    Dim logs(0 To 7) As String
    Dim errMsg As String
    Dim teardownError As String
    Dim pendingBefore As Scripting.Dictionary
    Dim pendingAfter As Scripting.Dictionary
    Dim expectedFixturePendingCount As Long

    logs(0) = "1. Arrange: ForceLocalBackend uses BackendSandbox"
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_CalidadCacheRiesgo_TareasRefresh_LimpiaDiccionarios = BuildFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Arrange: seed deterministic pending retipification graph"
    If Not SeedCalidadCacheGraph(True, expectedFixturePendingCount, errMsg) Then
        Test_CalidadCacheRiesgo_TareasRefresh_LimpiaDiccionarios = BuildFail(errMsg, logs)
        GoTo Cleanup
    End If

    logs(2) = "3. Arrange: stale cached pending task dictionary"
    Set pendingBefore = New Scripting.Dictionary
    pendingBefore.Add "STALE-RISK", "kept-before-refresh"
    Set m_ObjTareasCalidad = New TareasCalidad
    Set m_ObjTareasCalidad.ColRiesgosPorReTipificar = pendingBefore

    logs(3) = "4. Act: explicit TareasCalidad refresh entrypoint"
    CallByName m_ObjTareasCalidad, "RefrescarPendientesCalidad", VbMethod, errMsg

    logs(4) = "5. Assert: refresh propagates no error"
    If errMsg <> "" Then
        Test_CalidadCacheRiesgo_TareasRefresh_LimpiaDiccionarios = BuildFail(errMsg, logs)
        GoTo Cleanup
    End If

    logs(5) = "6. Assert: stale pending dictionary was cleared before recompute"
    Set pendingAfter = m_ObjTareasCalidad.ColRiesgosPorReTipificar
    If pendingAfter Is Nothing Then
        Test_CalidadCacheRiesgo_TareasRefresh_LimpiaDiccionarios = BuildFail("ColRiesgosPorReTipificar no puede ser Nothing tras refresco", logs)
        GoTo Cleanup
    End If
    If pendingAfter.Exists("STALE-RISK") Then
        Test_CalidadCacheRiesgo_TareasRefresh_LimpiaDiccionarios = BuildFail("RefrescarPendientesCalidad debe limpiar diccionarios pendientes cacheados", logs)
        GoTo Cleanup
    End If

    logs(6) = "7. Assert: fixture-owned pending cache keys are exact"
    If Not ValidatePendingFixtureKeys(pendingAfter, expectedFixturePendingCount, errMsg) Then
        Test_CalidadCacheRiesgo_TareasRefresh_LimpiaDiccionarios = BuildFail(errMsg, logs)
        GoTo Cleanup
    End If

    logs(7) = "8. Assert: deterministic fixture graph drove refresh"
    Test_CalidadCacheRiesgo_TareasRefresh_LimpiaDiccionarios = BuildOk("tareas_refresh_recomputes_fixture", logs)

Cleanup:
    On Error Resume Next
    Set pendingBefore = Nothing
    Set pendingAfter = Nothing
    teardownError = ""
    If Not TeardownCalidadCacheGraph(teardownError) Then
        Test_CalidadCacheRiesgo_TareasRefresh_LimpiaDiccionarios = ResultAfterTeardown(Test_CalidadCacheRiesgo_TareasRefresh_LimpiaDiccionarios, teardownError, logs)
    End If
    Test_Helper.ResetTestSession errMsg
    Exit Function
EH:
    Test_CalidadCacheRiesgo_TareasRefresh_LimpiaDiccionarios = BuildFail("Test_CalidadCacheRiesgo_TareasRefresh_LimpiaDiccionarios: " & Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_CalidadCacheRiesgo_RiesgoRegistrar_InvalidatesRiskCache() As String
    On Error GoTo EH

    Dim logs(0 To 8) As String
    Dim errMsg As String
    Dim teardownError As String
    Dim riesgoCacheAntes As riesgo
    Dim riesgoActual As riesgo
    Dim riesgoCacheDespues As riesgo
    Dim riesgoUnrelatedAntes As riesgo
    Dim riesgoUnrelatedDespues As riesgo
    Dim persistedDescription As String
    Dim expectedFixturePendingCount As Long

    logs(0) = "1. Arrange: ForceLocalBackend uses BackendSandbox"
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        Test_CalidadCacheRiesgo_RiesgoRegistrar_InvalidatesRiskCache = BuildFail(errMsg, logs)
        Exit Function
    End If

    logs(1) = "2. Arrange: seed deterministic normal risk fixture"
    If Not SeedCalidadCacheGraph(False, expectedFixturePendingCount, errMsg) Then
        Test_CalidadCacheRiesgo_RiesgoRegistrar_InvalidatesRiskCache = BuildFail(errMsg, logs)
        GoTo Cleanup
    End If

    logs(2) = "3. Arrange: cache target and unrelated risk references"
    Set riesgoCacheAntes = GetCachedRiesgo(CStr(FIX_RIESGO_ID), errMsg)
    If errMsg <> "" Then GoTo FailWithErr
    Set riesgoUnrelatedAntes = GetCachedRiesgo(CStr(FIX_RIESGO_UNRELATED_ID), errMsg)
    If errMsg <> "" Then GoTo FailWithErr
    Set riesgoActual = Constructor.getRiesgo(CStr(FIX_RIESGO_ID), , , errMsg)
    If errMsg <> "" Then GoTo FailWithErr
    If riesgoCacheAntes Is Nothing Or riesgoActual Is Nothing Or riesgoUnrelatedAntes Is Nothing Then
        Test_CalidadCacheRiesgo_RiesgoRegistrar_InvalidatesRiskCache = BuildFail("No se pudo preparar riesgos fixture", logs)
        GoTo Cleanup
    End If

    logs(3) = "4. Arrange: mutate a concrete persisted field"
    riesgoActual.Descripcion = FIX_UPDATED_DESCRIPTION

    logs(4) = "5. Act: Registrar should persist mutation and invalidate target risk cache"
    riesgoActual.Registrar EnumSiNo.No, riesgoCacheAntes, errMsg
    If errMsg <> "" Then GoTo FailWithErr

    logs(5) = "6. Assert: persisted mutation is stored in TbRiesgos"
    persistedDescription = RiskDescriptionFromDb(FIX_RIESGO_ID, errMsg)
    If errMsg <> "" Then GoTo FailWithErr
    If persistedDescription <> FIX_UPDATED_DESCRIPTION Then
        Test_CalidadCacheRiesgo_RiesgoRegistrar_InvalidatesRiskCache = BuildFail("Descripcion persistida no coincide", logs)
        GoTo Cleanup
    End If

    logs(6) = "7. Assert: cached target risk is replaced and has mutation"
    Set riesgoCacheDespues = GetCachedRiesgo(CStr(FIX_RIESGO_ID), errMsg)
    If errMsg <> "" Then GoTo FailWithErr
    If riesgoCacheDespues Is Nothing Or riesgoCacheDespues Is riesgoCacheAntes Then
        Test_CalidadCacheRiesgo_RiesgoRegistrar_InvalidatesRiskCache = BuildFail("Riesgo.Registrar debe invalidar la cache del riesgo guardado", logs)
        GoTo Cleanup
    End If
    If riesgoCacheDespues.Descripcion <> FIX_UPDATED_DESCRIPTION Then
        Test_CalidadCacheRiesgo_RiesgoRegistrar_InvalidatesRiskCache = BuildFail("La cache rehidratada debe contener la descripción actualizada", logs)
        GoTo Cleanup
    End If

    logs(7) = "8. Assert: unrelated cached risk is preserved"
    Set riesgoUnrelatedDespues = GetCachedRiesgo(CStr(FIX_RIESGO_UNRELATED_ID), errMsg)
    If errMsg <> "" Then GoTo FailWithErr
    If Not (riesgoUnrelatedDespues Is riesgoUnrelatedAntes) Then
        Test_CalidadCacheRiesgo_RiesgoRegistrar_InvalidatesRiskCache = BuildFail("Registrar no debe invalidar otro riesgo", logs)
        GoTo Cleanup
    End If

    logs(8) = "9. Assert: p_Error remained empty"
    Test_CalidadCacheRiesgo_RiesgoRegistrar_InvalidatesRiskCache = BuildOk("riesgo_registrar_invalidates_cache", logs)
    GoTo Cleanup

FailWithErr:
    Test_CalidadCacheRiesgo_RiesgoRegistrar_InvalidatesRiskCache = BuildFail(errMsg, logs)

Cleanup:
    On Error Resume Next
    teardownError = ""
    If Not TeardownCalidadCacheGraph(teardownError) Then
        Test_CalidadCacheRiesgo_RiesgoRegistrar_InvalidatesRiskCache = ResultAfterTeardown(Test_CalidadCacheRiesgo_RiesgoRegistrar_InvalidatesRiskCache, teardownError, logs)
    End If
    Test_Helper.ResetTestSession errMsg
    Exit Function
EH:
    Test_CalidadCacheRiesgo_RiesgoRegistrar_InvalidatesRiskCache = BuildFail("Test_CalidadCacheRiesgo_RiesgoRegistrar_InvalidatesRiskCache: " & Err.Description, logs)
    Resume Cleanup
End Function
