Attribute VB_Name = "Test_Workflow_Strict"
Option Compare Database
Option Explicit

' ============================================================================
' Test_Workflow_Strict — CAP-007 Workflow strict TDD v2.4.2 atoms (Fase B4, 2026-06-15)
'
' Touched table: tbLogEstados. Tests at the repository level only.
' EjecutarTransicion and ReabrirSolicitudCerrada require non-trivial setup
' (rol-aware m_ObjUsuarioActivo, multi-table preconditions, transactional
' workspaces) and are documented as deferred — see strict-tdd-coverage §5.
' LogEstadoRepositorio is the public seam: it accepts db injection and
' does not destroy the caller's connection.
'
' Fixture IDs: idSolicitud 900911-900999.
' ============================================================================

Private Const ID_SOL_INSERT As Long = 900911
Private Const ID_SOL_HISTORIAL As Long = 900912

' ----------------------------------------------------------------------------
' Setup / teardown helpers
' ----------------------------------------------------------------------------

Private Sub SetupWorkflowSandbox(ByRef p_TempRoot As String, ByRef p_Error As String)
    On Error GoTo EH
    Dim runError As String
    Dim logs() As String
    logs = TestHelper.NewLogsArray(3)

    p_Error = ""
    p_TempRoot = ""

    ' v2.4.3 harness: hardened production guard + sandbox URL validation
    ' moved to TestHelper.BeginTestSession (UNC, fingerprint, FSO+DAO).
    ' SetupProdGlobalsForTest covers the 4 prod globals + tempRoot.
    If Not TestHelper.BeginTestSession(logs, runError) Then
        p_Error = runError
        Exit Sub
    End If

    Call TestHelper.SetupProdGlobalsForTest(runError, p_TempRoot, "cap007_wf", "CAP007 Test User")
    If runError <> "" Then
        p_Error = runError
        Call TestHelper.ResetTestSession
        Exit Sub
    End If

    Exit Sub

EH:
    p_Error = "TESTS BLOCKED: SetupWorkflowSandbox failed: " & Err.Number & " - " & Err.Description
    Call TestHelper.ResetTestSession
End Sub

Private Sub TeardownFixtures(ByVal p_Db As DAO.Database)
    On Error Resume Next
    p_Db.Execute "DELETE FROM tbLogEstados WHERE idSolicitud >= 900910 AND idSolicitud <= 900999", dbFailOnError
    p_Db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud >= 900910 AND idSolicitud <= 900999", dbFailOnError
    p_Db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente >= 900910 AND IDExpediente <= 900999", dbFailOnError
    On Error GoTo 0
End Sub

Private Sub SeedSolicitud(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long, ByVal p_Codigo As String)
    On Error Resume Next
    p_Db.Execute "DELETE FROM tbLogEstados WHERE idSolicitud=" & p_IdSolicitud, dbFailOnError
    p_Db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud=" & p_IdSolicitud, dbFailOnError
    p_Db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & p_IdSolicitud, dbFailOnError
    On Error GoTo 0
    p_Db.Execute "INSERT INTO TbExpedientes (IDExpediente) VALUES (" & p_IdSolicitud & ")", dbFailOnError
    Dim sql As String
    sql = "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, " & _
          "idEstadoInterno, fechaCreacion, usuarioCreacion, revisionCalidadEstado) VALUES (" & _
          p_IdSolicitud & ", " & p_IdSolicitud & ", 'PC', '" & p_Codigo & "', 2, Now(), " & _
          "'TestWorkflowStrict', 'PENDIENTE')"
    p_Db.Execute sql, dbFailOnError
End Sub

Private Sub BuildLogEstado(ByRef p_Log As LogEstado, ByVal p_IdSolicitud As Long, _
                            ByVal p_Anterior As Long, ByVal p_Nuevo As Long, _
                            ByVal p_Fecha As Date, ByVal p_Usuario As String)
    Set p_Log = New LogEstado
    p_Log.idLogEstado = 0
    p_Log.idSolicitud = p_IdSolicitud
    p_Log.idEstadoAnterior = p_Anterior
    p_Log.idEstadoNuevo = p_Nuevo
    p_Log.fechaTransicion = p_Fecha
    p_Log.usuarioTransicion = p_Usuario
End Sub

Private Function CountLogEstados(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long) As Long
    Dim rs As DAO.Recordset
    On Error GoTo CleanExit
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS n FROM tbLogEstados WHERE idSolicitud=" & p_IdSolicitud, dbOpenSnapshot)
    If Not rs.EOF Then CountLogEstados = Nz(rs!n, 0)
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

Private Function GetLastLogEstadoRow(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long) As DAO.Recordset
    Dim rs As DAO.Recordset
    Set rs = p_Db.OpenRecordset( _
        "SELECT TOP 1 idEstadoAnterior, idEstadoNuevo, fechaTransicion, usuarioTransicion " & _
        "FROM tbLogEstados WHERE idSolicitud=" & p_IdSolicitud & " " & _
        "ORDER BY fechaTransicion DESC, idLogEstado DESC", dbOpenSnapshot)
    Set GetLastLogEstadoRow = rs
End Function

' ----------------------------------------------------------------------------
' Atoms
' ----------------------------------------------------------------------------

Public Function Test_Workflow_Strict_LogEstadoRepositorio_Guardar_InsertsRowWithExpectedFields() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim setupError As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim logEstado As LogEstado
    Dim countBefore As Long
    Dim countAfter As Long
    Dim fechaTest As Date
    Dim rs As DAO.Recordset

    Call SetupWorkflowSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Workflow_Strict_LogEstadoRepositorio_Guardar_InsertsRowWithExpectedFields = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_INSERT, "CAP007-INSERT")

    countBefore = CountLogEstados(db, ID_SOL_INSERT)
    logs(0) = "1. countBefore=" & countBefore

    fechaTest = #6/15/2026 10:00:00 AM#
    Call BuildLogEstado(logEstado, ID_SOL_INSERT, 2, 3, fechaTest, "workflow-test-user")
    LogEstadoRepositorio.Guardar logEstado, db
    logs(1) = "2. LogEstadoRepositorio.Guardar called with explicit DAO.Database"

    countAfter = CountLogEstados(db, ID_SOL_INSERT)
    If countAfter <> countBefore + 1 Then
        logs(2) = "3. FAILED: cardinality wrong: before=" & countBefore & " after=" & countAfter
        Test_Workflow_Strict_LogEstadoRepositorio_Guardar_InsertsRowWithExpectedFields = TestHelper.BuildJsonFail("cardinalidad wrong", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. countAfter=" & countAfter

    Set rs = db.OpenRecordset( _
        "SELECT idEstadoAnterior, idEstadoNuevo, fechaTransicion, usuarioTransicion " & _
        "FROM tbLogEstados WHERE idSolicitud=" & ID_SOL_INSERT, dbOpenSnapshot)
    If rs.EOF Then
        logs(3) = "4. FAILED: row not found"
        Test_Workflow_Strict_LogEstadoRepositorio_Guardar_InsertsRowWithExpectedFields = TestHelper.BuildJsonFail("row not found", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. row found: anterior=" & Nz(rs!idEstadoAnterior, 0) & " nuevo=" & Nz(rs!idEstadoNuevo, 0) & " usuario=[" & Nz(rs!usuarioTransicion, "") & "]"

    If Nz(rs!idEstadoAnterior, 0) <> 2 Then
        logs(4) = "5. FAILED: idEstadoAnterior mismatch: " & Nz(rs!idEstadoAnterior, 0)
        Test_Workflow_Strict_LogEstadoRepositorio_Guardar_InsertsRowWithExpectedFields = TestHelper.BuildJsonFail("idEstadoAnterior mismatch", logs)
        GoTo CleanExit
    End If
    If Nz(rs!idEstadoNuevo, 0) <> 3 Then
        logs(4) = "5. FAILED: idEstadoNuevo mismatch: " & Nz(rs!idEstadoNuevo, 0)
        Test_Workflow_Strict_LogEstadoRepositorio_Guardar_InsertsRowWithExpectedFields = TestHelper.BuildJsonFail("idEstadoNuevo mismatch", logs)
        GoTo CleanExit
    End If
    If Nz(rs!usuarioTransicion, "") <> "workflow-test-user" Then
        logs(4) = "5. FAILED: usuarioTransicion mismatch: [" & Nz(rs!usuarioTransicion, "") & "]"
        Test_Workflow_Strict_LogEstadoRepositorio_Guardar_InsertsRowWithExpectedFields = TestHelper.BuildJsonFail("usuarioTransicion mismatch", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. idEstadoAnterior, idEstadoNuevo, usuarioTransicion persisted correctly"

    logs(5) = "6. PASS"
    Test_Workflow_Strict_LogEstadoRepositorio_Guardar_InsertsRowWithExpectedFields = TestHelper.BuildJsonOk("true", logs)
    Exit Function

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Workflow_Strict_LogEstadoRepositorio_Guardar_InsertsRowWithExpectedFields = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Workflow_Strict_LogEstadoRepositorio_Guardar_AutoAssignsIdLogEstado() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim logEstado As LogEstado
    Dim idBefore As Long
    Dim rs As DAO.Recordset

    Call SetupWorkflowSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Workflow_Strict_LogEstadoRepositorio_Guardar_AutoAssignsIdLogEstado = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_INSERT, "CAP007-AUTOID")

    Call BuildLogEstado(logEstado, ID_SOL_INSERT, 2, 3, Now(), "auto-id-test")
    idBefore = logEstado.idLogEstado
    logs(0) = "1. idLogEstado before Guardar=" & idBefore

    LogEstadoRepositorio.Guardar logEstado, db
    logs(1) = "2. Guardar executed"

    Set rs = db.OpenRecordset( _
        "SELECT TOP 1 idLogEstado FROM tbLogEstados WHERE idSolicitud=" & ID_SOL_INSERT & " " & _
        "ORDER BY idLogEstado DESC", dbOpenSnapshot)
    If rs.EOF Then
        logs(2) = "3. FAILED: no row inserted"
        Test_Workflow_Strict_LogEstadoRepositorio_Guardar_AutoAssignsIdLogEstado = TestHelper.BuildJsonFail("no row inserted", logs)
        GoTo CleanExit
    End If

    If Nz(rs!idLogEstado, 0) <= 0 Then
        logs(2) = "3. FAILED: autonumerico idLogEstado=" & Nz(rs!idLogEstado, 0) & " expected > 0"
        Test_Workflow_Strict_LogEstadoRepositorio_Guardar_AutoAssignsIdLogEstado = TestHelper.BuildJsonFail("idLogEstado not assigned", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. persisted idLogEstado=" & Nz(rs!idLogEstado, 0)

    logs(3) = "4. PASS"
    Test_Workflow_Strict_LogEstadoRepositorio_Guardar_AutoAssignsIdLogEstado = TestHelper.BuildJsonOk("true", logs)
    Exit Function

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Workflow_Strict_LogEstadoRepositorio_Guardar_AutoAssignsIdLogEstado = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Workflow_Strict_LogEstadoRepositorio_getUltimoEstadoAnterior_ReturnsLastTransition() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(7)
    Dim setupError As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim logEstado1 As LogEstado
    Dim logEstado2 As LogEstado
    Dim logEstado3 As LogEstado
    Dim result As Long

    Call SetupWorkflowSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Workflow_Strict_LogEstadoRepositorio_getUltimoEstadoAnterior_ReturnsLastTransition = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_HISTORIAL, "CAP007-HIST")

    ' Insertar 3 transiciones con fechas estrictamente crecientes para
    ' asegurar que ORDER BY fechaTransicion DESC distingue la ultima.
    Call BuildLogEstado(logEstado1, ID_SOL_HISTORIAL, 2, 3, #6/15/2026 09:00:00 AM#, "step1")
    Call BuildLogEstado(logEstado2, ID_SOL_HISTORIAL, 3, 4, #6/15/2026 10:00:00 AM#, "step2")
    Call BuildLogEstado(logEstado3, ID_SOL_HISTORIAL, 4, 5, #6/15/2026 11:00:00 AM#, "step3")
    LogEstadoRepositorio.Guardar logEstado1, db
    LogEstadoRepositorio.Guardar logEstado2, db
    LogEstadoRepositorio.Guardar logEstado3, db
    logs(0) = "1. Inserted 3 transiciones: 2->3 (09:00), 3->4 (10:00), 4->5 (11:00)"

    result = LogEstadoRepositorio.getUltimoEstadoAnterior(ID_SOL_HISTORIAL, db)
    logs(1) = "2. getUltimoEstadoAnterior returned: " & result

    If result <> 4 Then
        logs(2) = "3. FAILED: expected 4 (anterior de la ultima 4->5), got " & result
        Test_Workflow_Strict_LogEstadoRepositorio_getUltimoEstadoAnterior_ReturnsLastTransition = TestHelper.BuildJsonFail("expected 4", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. getUltimoEstadoAnterior=4 as expected"

    ' Verificar que devuelve 0 para una solicitud sin historial
    result = LogEstadoRepositorio.getUltimoEstadoAnterior(999999, db)
    logs(3) = "4. getUltimoEstadoAnterior(999999) returned: " & result

    If result <> 0 Then
        logs(4) = "5. FAILED: expected 0 (no history), got " & result
        Test_Workflow_Strict_LogEstadoRepositorio_getUltimoEstadoAnterior_ReturnsLastTransition = TestHelper.BuildJsonFail("expected 0 for empty", logs)
        GoTo CleanExit
    End If
    logs(4) = "5. getUltimoEstadoAnterior=0 as expected for empty history"

    logs(5) = "6. PASS"
    Test_Workflow_Strict_LogEstadoRepositorio_getUltimoEstadoAnterior_ReturnsLastTransition = TestHelper.BuildJsonOk("true", logs)
    Exit Function

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Workflow_Strict_LogEstadoRepositorio_getUltimoEstadoAnterior_ReturnsLastTransition = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Workflow_Strict_LogEstadoRepositorio_getHistorialPorIdSolicitud_ReturnsOrderedHistory() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim setupError As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim logEstado1 As LogEstado
    Dim logEstado2 As LogEstado
    Dim logEstado3 As LogEstado
    Dim hist As Scripting.Dictionary
    Dim firstKey As Variant
    Dim firstEstadoNuevo As Long
    Dim lastKey As Variant
    Dim lastEstadoNuevo As Long

    Call SetupWorkflowSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Workflow_Strict_LogEstadoRepositorio_getHistorialPorIdSolicitud_ReturnsOrderedHistory = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_HISTORIAL, "CAP007-HIST2")

    Call BuildLogEstado(logEstado1, ID_SOL_HISTORIAL, 2, 3, #6/15/2026 09:00:00 AM#, "step1")
    Call BuildLogEstado(logEstado2, ID_SOL_HISTORIAL, 3, 4, #6/15/2026 10:00:00 AM#, "step2")
    Call BuildLogEstado(logEstado3, ID_SOL_HISTORIAL, 4, 5, #6/15/2026 11:00:00 AM#, "step3")
    LogEstadoRepositorio.Guardar logEstado1, db
    LogEstadoRepositorio.Guardar logEstado2, db
    LogEstadoRepositorio.Guardar logEstado3, db
    logs(0) = "1. Inserted 3 transiciones (2->3 09:00, 3->4 10:00, 4->5 11:00)"

    Set hist = LogEstadoRepositorio.getHistorialPorIdSolicitud(ID_SOL_HISTORIAL, db)
    If hist Is Nothing Then
        logs(1) = "2. FAILED: getHistorial returned Nothing"
        Test_Workflow_Strict_LogEstadoRepositorio_getHistorialPorIdSolicitud_ReturnsOrderedHistory = TestHelper.BuildJsonFail("getHistorial returned Nothing", logs)
        GoTo CleanExit
    End If

    If hist.Count <> 3 Then
        logs(1) = "2. FAILED: expected Count=3, got " & hist.Count
        Test_Workflow_Strict_LogEstadoRepositorio_getHistorialPorIdSolicitud_ReturnsOrderedHistory = TestHelper.BuildJsonFail("Count mismatch", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. Dictionary Count=3 as expected"

    firstKey = hist.Keys(0)
    Set firstKey = hist(firstKey)
    firstEstadoNuevo = firstKey.idEstadoNuevo
    logs(2) = "3. first key (oldest): idEstadoNuevo=" & firstEstadoNuevo

    If firstEstadoNuevo <> 3 Then
        logs(3) = "4. FAILED: expected first.idEstadoNuevo=3 (oldest: 2->3), got " & firstEstadoNuevo
        Test_Workflow_Strict_LogEstadoRepositorio_getHistorialPorIdSolicitud_ReturnsOrderedHistory = TestHelper.BuildJsonFail("first mismatch", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. first.idEstadoNuevo=3 (oldest transition) as expected"

    lastKey = hist.Keys(hist.Count - 1)
    Set lastKey = hist(lastKey)
    lastEstadoNuevo = lastKey.idEstadoNuevo
    logs(4) = "5. last key (newest): idEstadoNuevo=" & lastEstadoNuevo

    If lastEstadoNuevo <> 5 Then
        logs(5) = "6. FAILED: expected last.idEstadoNuevo=5 (newest: 4->5), got " & lastEstadoNuevo
        Test_Workflow_Strict_LogEstadoRepositorio_getHistorialPorIdSolicitud_ReturnsOrderedHistory = TestHelper.BuildJsonFail("last mismatch", logs)
        GoTo CleanExit
    End If
    logs(5) = "6. last.idEstadoNuevo=5 (newest transition) as expected — history ordered ASC by fechaTransicion"

    Test_Workflow_Strict_LogEstadoRepositorio_getHistorialPorIdSolicitud_ReturnsOrderedHistory = TestHelper.BuildJsonOk("true", logs)
    Exit Function

CleanExit:
    On Error Resume Next
    Set hist = Nothing
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Workflow_Strict_LogEstadoRepositorio_getHistorialPorIdSolicitud_ReturnsOrderedHistory = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Workflow_Strict_LogEstadoRepositorio_EliminarPorIdSolicitud_DeletesAllRowsForSolicitud() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim logEstado1 As LogEstado
    Dim logEstado2 As LogEstado
    Dim countBefore As Long
    Dim countAfter As Long

    Call SetupWorkflowSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Workflow_Strict_LogEstadoRepositorio_EliminarPorIdSolicitud_DeletesAllRowsForSolicitud = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_HISTORIAL, "CAP007-DEL")

    Call BuildLogEstado(logEstado1, ID_SOL_HISTORIAL, 2, 3, #6/15/2026 09:00:00 AM#, "del1")
    Call BuildLogEstado(logEstado2, ID_SOL_HISTORIAL, 3, 4, #6/15/2026 10:00:00 AM#, "del2")
    LogEstadoRepositorio.Guardar logEstado1, db
    LogEstadoRepositorio.Guardar logEstado2, db
    logs(0) = "1. Inserted 2 transiciones for ID_SOL_HISTORIAL"

    countBefore = CountLogEstados(db, ID_SOL_HISTORIAL)
    If countBefore <> 2 Then
        logs(1) = "2. FAILED: seed cardinality expected 2, got " & countBefore
        Test_Workflow_Strict_LogEstadoRepositorio_EliminarPorIdSolicitud_DeletesAllRowsForSolicitud = TestHelper.BuildJsonFail("seed cardinalidad", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. countBefore=2"

    LogEstadoRepositorio.EliminarPorIdSolicitud ID_SOL_HISTORIAL, db
    logs(2) = "3. EliminarPorIdSolicitud called with explicit DAO.Database"

    countAfter = CountLogEstados(db, ID_SOL_HISTORIAL)
    If countAfter <> 0 Then
        logs(3) = "4. FAILED: expected count=0 after delete, got " & countAfter
        Test_Workflow_Strict_LogEstadoRepositorio_EliminarPorIdSolicitud_DeletesAllRowsForSolicitud = TestHelper.BuildJsonFail("delete cardinalidad", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. countAfter=0 as expected (all rows for ID_SOL_HISTORIAL deleted)"

    Test_Workflow_Strict_LogEstadoRepositorio_EliminarPorIdSolicitud_DeletesAllRowsForSolicitud = TestHelper.BuildJsonOk("true", logs)
    Exit Function

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Workflow_Strict_LogEstadoRepositorio_EliminarPorIdSolicitud_DeletesAllRowsForSolicitud = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Workflow_Strict_ReabrirSolicitudCerrada_RaisesErrorWhenIdEstadoNotAprobada() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim wfServ As New WorkflowServicio
    Dim usuario As Usuario
    Dim sol As Solicitud
    Dim errorNumber As Long
    Dim errorDescription As String

    Call SetupWorkflowSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Workflow_Strict_ReabrirSolicitudCerrada_RaisesErrorWhenIdEstadoNotAprobada = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_HISTORIAL, "CAP007-REABRIR-SAD")

    Set usuario = New Usuario
    usuario.nombre = "ReabrirTestUser"
    usuario.rol = rol.Calidad

    Set sol = New Solicitud
    sol.idSolicitud = ID_SOL_HISTORIAL
    sol.idEstadoInterno = 7  ' estadoFormalizacion, NOT estadoAprobada (8)

    On Error Resume Next
    wfServ.ReabrirSolicitudCerrada sol, usuario, db
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error GoTo EH
    logs(0) = "1. Called ReabrirSolicitudCerrada with idEstadoInterno=7 (estadoFormalizacion)"

    If errorNumber = 0 Then
        logs(1) = "2. FAILED: expected CondorError Raise (513), but no error fired"
        Test_Workflow_Strict_ReabrirSolicitudCerrada_RaisesErrorWhenIdEstadoNotAprobada = TestHelper.BuildJsonFail("expected Raise", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. Caught error " & errorNumber & ": " & errorDescription

    If errorNumber <> 513 Then
        logs(2) = "3. FAILED: expected errorNumber=513, got " & errorNumber
        Test_Workflow_Strict_ReabrirSolicitudCerrada_RaisesErrorWhenIdEstadoNotAprobada = TestHelper.BuildJsonFail("expected 513", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. errorNumber=513 as expected (custom application error: la solicitud no se encuentra en estado APROBADA)"

    logs(3) = "4. PASS"
    Test_Workflow_Strict_ReabrirSolicitudCerrada_RaisesErrorWhenIdEstadoNotAprobada = TestHelper.BuildJsonOk("true", logs)
    Exit Function

CleanExit:
    On Error Resume Next
    Set usuario = Nothing
    Set sol = Nothing
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Workflow_Strict_ReabrirSolicitudCerrada_RaisesErrorWhenIdEstadoNotAprobada = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Workflow_Strict_ReabrirSolicitudCerrada_RaisesErrorWhenRolIsTecnico() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim wfServ As New WorkflowServicio
    Dim usuario As Usuario
    Dim sol As Solicitud
    Dim errorNumber As Long
    Dim errorDescription As String

    Call SetupWorkflowSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Workflow_Strict_ReabrirSolicitudCerrada_RaisesErrorWhenRolIsTecnico = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_HISTORIAL, "CAP007-REABRIR-ROLTEC")

    Set usuario = New Usuario
    usuario.nombre = "ReabrirTecnicoTest"
    usuario.rol = rol.Tecnico

    Set sol = New Solicitud
    sol.idSolicitud = ID_SOL_HISTORIAL
    sol.idEstadoInterno = 8  ' estadoAprobada (correct state)

    On Error Resume Next
    wfServ.ReabrirSolicitudCerrada sol, usuario, db
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error GoTo EH
    logs(0) = "1. Called ReabrirSolicitudCerrada with usuario.rol=Tecnico"

    If errorNumber = 0 Then
        logs(1) = "2. FAILED: expected CondorError Raise (513), but no error fired"
        Test_Workflow_Strict_ReabrirSolicitudCerrada_RaisesErrorWhenRolIsTecnico = TestHelper.BuildJsonFail("expected Raise", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. Caught error " & errorNumber & ": " & errorDescription

    If errorNumber <> 513 Then
        logs(2) = "3. FAILED: expected errorNumber=513, got " & errorNumber
        Test_Workflow_Strict_ReabrirSolicitudCerrada_RaisesErrorWhenRolIsTecnico = TestHelper.BuildJsonFail("expected 513", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. errorNumber=513 as expected (custom application error: no tiene permisos para reabrir una solicitud)"

    logs(3) = "4. PASS"
    Test_Workflow_Strict_ReabrirSolicitudCerrada_RaisesErrorWhenRolIsTecnico = TestHelper.BuildJsonOk("true", logs)
    Exit Function

CleanExit:
    On Error Resume Next
    Set usuario = Nothing
    Set sol = Nothing
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Workflow_Strict_ReabrirSolicitudCerrada_RaisesErrorWhenRolIsTecnico = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Workflow_Strict_RunAll() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim result As String

    result = Test_Workflow_Strict_LogEstadoRepositorio_Guardar_InsertsRowWithExpectedFields()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(0) = "1. LogEstadoRepositorio.Guardar inserts row with expected fields"

    result = Test_Workflow_Strict_LogEstadoRepositorio_Guardar_AutoAssignsIdLogEstado()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(1) = "2. LogEstadoRepositorio.Guardar assigns autonumerico idLogEstado"

    result = Test_Workflow_Strict_LogEstadoRepositorio_getUltimoEstadoAnterior_ReturnsLastTransition()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(2) = "3. getUltimoEstadoAnterior returns the correct last transition (and 0 for empty history)"

    result = Test_Workflow_Strict_LogEstadoRepositorio_getHistorialPorIdSolicitud_ReturnsOrderedHistory()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(3) = "4. getHistorialPorIdSolicitud returns the full ordered history (ASC by fechaTransicion)"

    result = Test_Workflow_Strict_LogEstadoRepositorio_EliminarPorIdSolicitud_DeletesAllRowsForSolicitud()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(4) = "5. EliminarPorIdSolicitud deletes all rows for the solicitud"

    result = Test_Workflow_Strict_ReabrirSolicitudCerrada_RaisesErrorWhenIdEstadoNotAprobada()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(5) = "6. ReabrirSolicitudCerrada raises 513 when idEstadoInterno != estadoAprobada"

    result = Test_Workflow_Strict_ReabrirSolicitudCerrada_RaisesErrorWhenRolIsTecnico()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(6) = "7. ReabrirSolicitudCerrada raises 513 when usuario.rol=Tecnico"

    logs(7) = "8. PASS"
    Test_Workflow_Strict_RunAll = TestHelper.BuildJsonOk("true", logs)
    Exit Function

Failed:
    Test_Workflow_Strict_RunAll = TestHelper.BuildJsonFail("one CAP-007 Workflow atom failed", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Workflow_Strict_RunAll: ERR " & Err.Number & " - " & Err.Description
    Test_Workflow_Strict_RunAll = TestHelper.BuildJsonFail(Err.Description, logs)
End Function
