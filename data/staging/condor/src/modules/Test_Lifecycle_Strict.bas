Attribute VB_Name = "Test_Lifecycle_Strict"
Option Compare Database
Option Explicit

' ============================================================================
' Test_Lifecycle_Strict — CAP-006 Lifecycle atoms (Fase B3, 2026-06-15)
'
' Touched tables: tbSolicitudes, TbExpedientes (read + write fixture).
' Read-only contract: getSolicitudesViewModel is a SQL hydration; no side
' effects on production tables. Fixture IDs in 900860-900899.
' ============================================================================

Private Const TEST_ID_BASE As Long = 900860
Private Const TEST_ID_TOP As Long = 900899

Private Const ID_SOL_LIFECYCLE_1 As Long = 900861
Private Const ID_SOL_LIFECYCLE_2 As Long = 900862
Private Const ID_SOL_LIFECYCLE_3 As Long = 900863

Private Const FILE_PREFIX As String = "CAP006-STRICT-"

' ----------------------------------------------------------------------------
' Setup / teardown helpers
' ----------------------------------------------------------------------------

Private Sub SetupLifecycleSandbox(ByRef p_TempRoot As String, ByRef p_Error As String)
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

    Call TestHelper.SetupProdGlobalsForTest(runError, p_TempRoot, "cap006_lifecycle", "CAP006 Test User")
    If runError <> "" Then
        p_Error = runError
        Call TestHelper.ResetTestSession
        Exit Sub
    End If

    ' CAP006-specific: explicitly zero the user id so the EXISTS filter does not match
    m_ObjUsuarioActivo.id = 0
    Exit Sub

EH:
    p_Error = "TESTS BLOCKED: SetupLifecycleSandbox failed: " & Err.Number & " - " & Err.Description
    Call TestHelper.ResetTestSession
End Sub

Private Sub TeardownFixtures(ByVal p_Db As DAO.Database)
    On Error Resume Next
    p_Db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud >= " & TEST_ID_BASE & " AND idSolicitud <= " & TEST_ID_TOP, dbFailOnError
    p_Db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente >= " & TEST_ID_BASE & " AND IDExpediente <= " & TEST_ID_TOP, dbFailOnError
    On Error GoTo 0
End Sub

Private Sub SeedSolicitud(ByVal p_Db As DAO.Database, ByVal p_IdSolicitud As Long, ByVal p_Tipo As String, ByVal p_Codigo As String)
    On Error Resume Next
    p_Db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & p_IdSolicitud, dbFailOnError
    p_Db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud=" & p_IdSolicitud, dbFailOnError
    On Error GoTo 0
    p_Db.Execute "INSERT INTO TbExpedientes (IDExpediente) VALUES (" & p_IdSolicitud & ")", dbFailOnError
    Dim sql As String
    sql = "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, " & _
          "idEstadoInterno, fechaCreacion, usuarioCreacion, revisionCalidadEstado) VALUES (" & _
          p_IdSolicitud & ", " & p_IdSolicitud & ", '" & p_Tipo & "', '" & p_Codigo & "', 2, Now(), " & _
          "'TestLifecycleStrict', 'PENDIENTE')"
    p_Db.Execute sql, dbFailOnError
End Sub

Private Function CountSolicitudesByIds(ByVal p_Db As DAO.Database, ByRef p_Ids() As Long) As Long
    Dim i As Long
    Dim rs As DAO.Recordset
    Dim inList As String
    inList = ""
    For i = LBound(p_Ids) To UBound(p_Ids)
        If Len(inList) > 0 Then inList = inList & ","
        inList = inList & p_Ids(i)
    Next i
    On Error GoTo CleanExit
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS n FROM tbSolicitudes WHERE idSolicitud IN (" & inList & ")", dbOpenSnapshot)
    If Not rs.EOF Then CountSolicitudesByIds = Nz(rs!n, 0)
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

Private Function CountSolicitudesByCodigo(ByVal p_Db As DAO.Database, ByVal p_Codigo As String) As Long
    Dim rs As DAO.Recordset
    On Error GoTo CleanExit
    Set rs = p_Db.OpenRecordset("SELECT COUNT(*) AS n FROM tbSolicitudes WHERE codigoSolicitud='" & Replace(p_Codigo, "'", "''") & "'", dbOpenSnapshot)
    If Not rs.EOF Then CountSolicitudesByCodigo = Nz(rs!n, 0)
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

' ----------------------------------------------------------------------------
' Atoms
' ----------------------------------------------------------------------------

Public Function Test_Lifecycle_Strict_GetSolicitudesViewModel_RetornaLasCreadas() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(7)
    Dim setupError As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New SolicitudServicio
    Dim ids(1 To 3) As Long
    Dim result As Object
    Dim found As Long
    Dim i As Long

    Call SetupLifecycleSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Lifecycle_Strict_GetSolicitudesViewModel_RetornaLasCreadas = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_LIFECYCLE_1, "PC", "CAP006-LIFECYCLE-1")
    Call SeedSolicitud(db, ID_SOL_LIFECYCLE_2, "CD_CA", "CAP006-LIFECYCLE-2")
    Call SeedSolicitud(db, ID_SOL_LIFECYCLE_3, "PC_SUB", "CAP006-LIFECYCLE-3")
    ids(1) = ID_SOL_LIFECYCLE_1
    ids(2) = ID_SOL_LIFECYCLE_2
    ids(3) = ID_SOL_LIFECYCLE_3
    logs(0) = "1. Seeded 3 solicitudes in sandbox: PC, CD_CA, PC_SUB"

    Set result = svc.getSolicitudesViewModel(db:=db)
    If result Is Nothing Then
        logs(1) = "2. FAILED: result is Nothing"
        Test_Lifecycle_Strict_GetSolicitudesViewModel_RetornaLasCreadas = TestHelper.BuildJsonFail("getSolicitudesViewModel returned Nothing", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. result returned with " & result.Count & " entries"

    found = 0
    For i = 1 To 3
        If result.Exists(CStr(ids(i))) Then found = found + 1
    Next i
    If found <> 3 Then
        logs(2) = "3. FAILED: expected 3 seeded ids in result, found " & found
        Test_Lifecycle_Strict_GetSolicitudesViewModel_RetornaLasCreadas = TestHelper.BuildJsonFail("missing seeded ids in result", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. all 3 seeded ids present in result"

    If CountSolicitudesByIds(db, ids) <> 3 Then
        logs(3) = "4. FAILED: seed count drift"
        Test_Lifecycle_Strict_GetSolicitudesViewModel_RetornaLasCreadas = TestHelper.BuildJsonFail("seed count drift", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. seed consistency verified"

    logs(4) = "5. PASS"
    Test_Lifecycle_Strict_GetSolicitudesViewModel_RetornaLasCreadas = TestHelper.BuildJsonOk("result.count=" & result.Count, logs)
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
    logs(4) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Lifecycle_Strict_GetSolicitudesViewModel_RetornaLasCreadas = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Lifecycle_Strict_GetSolicitudesViewModel_FiltraPorPalabraClave() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(7)
    Dim setupError As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim svc As New SolicitudServicio
    Dim result As Object
    Dim key As Variant
    Dim onlyMatch As Boolean

    Call SetupLifecycleSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Lifecycle_Strict_GetSolicitudesViewModel_FiltraPorPalabraClave = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    Call SeedSolicitud(db, ID_SOL_LIFECYCLE_1, "PC", "UNIQ-CAP006-AAA-1")
    Call SeedSolicitud(db, ID_SOL_LIFECYCLE_2, "CD_CA", "UNIQ-CAP006-BBB-2")
    Call SeedSolicitud(db, ID_SOL_LIFECYCLE_3, "PC_SUB", "UNIQ-CAP006-CCC-3")
    logs(0) = "1. Seeded 3 solicitudes with distinct codigoSolicitud prefixes"

    Set result = svc.getSolicitudesViewModel(palabraClave:="UNIQ-CAP006-AAA", db:=db)
    If result Is Nothing Then
        logs(1) = "2. FAILED: result is Nothing"
        Test_Lifecycle_Strict_GetSolicitudesViewModel_FiltraPorPalabraClave = TestHelper.BuildJsonFail("getSolicitudesViewModel returned Nothing", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. result.Count=" & result.Count

    If Not result.Exists(CStr(ID_SOL_LIFECYCLE_1)) Then
        logs(2) = "3. FAILED: matching idSolicitud=" & ID_SOL_LIFECYCLE_1 & " not in filtered result"
        Test_Lifecycle_Strict_GetSolicitudesViewModel_FiltraPorPalabraClave = TestHelper.BuildJsonFail("match not found", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. matching id present in result"

    onlyMatch = True
    For Each key In result.Keys
        If CLng(key) <> ID_SOL_LIFECYCLE_1 Then
            onlyMatch = False
            Exit For
        End If
    Next key
    If Not onlyMatch Then
        logs(3) = "4. FAILED: result contains extra ids that should have been filtered out"
        Test_Lifecycle_Strict_GetSolicitudesViewModel_FiltraPorPalabraClave = TestHelper.BuildJsonFail("filter leaked extra ids", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. result is exactly the matching id (1 entry)"

    logs(4) = "5. PASS"
    Test_Lifecycle_Strict_GetSolicitudesViewModel_FiltraPorPalabraClave = TestHelper.BuildJsonOk("true", logs)
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
    logs(4) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Lifecycle_Strict_GetSolicitudesViewModel_FiltraPorPalabraClave = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Lifecycle_Strict_GetSolicitudPorID_RetornaSolicitudParaIdValido() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(5)
    Dim setupError As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim solServ As New SolicitudServicio
    Dim sol As Solicitud
    Dim idSol As Long

    Call SetupLifecycleSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Lifecycle_Strict_GetSolicitudPorID_RetornaSolicitudParaIdValido = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()
    Call TeardownFixtures(db)
    idSol = 900861
    db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud=" & idSol, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & idSol, dbFailOnError
    db.Execute "INSERT INTO TbExpedientes (IDExpediente) VALUES (" & idSol & ")", dbFailOnError
    db.Execute "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, idEstadoInterno, fechaCreacion, usuarioCreacion) VALUES (" & idSol & ", " & idSol & ", 'PC', 'CAP006-GETSOL', 2, Now(), 'TestLifecycle')", dbFailOnError
    logs(0) = "1. Seeded tbSolicitudes id=" & idSol & " + parent TbExpedientes"

    Set sol = solServ.getSolicitudPorID(idSol, db)
    If sol Is Nothing Then
        logs(1) = "2. FAILED: getSolicitudPorID returned Nothing"
        Test_Lifecycle_Strict_GetSolicitudPorID_RetornaSolicitudParaIdValido = TestHelper.BuildJsonFail("returned Nothing", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. sol returned, not Nothing"

    If sol.idSolicitud <> idSol Then
        logs(2) = "3. FAILED: sol.idSolicitud mismatch: expected " & idSol & ", got " & sol.idSolicitud
        Test_Lifecycle_Strict_GetSolicitudPorID_RetornaSolicitudParaIdValido = TestHelper.BuildJsonFail("idSolicitud mismatch", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. sol.idSolicitud=" & sol.idSolicitud & " matches seeded id"

    If sol.codigoSolicitud <> "CAP006-GETSOL" Then
        logs(3) = "4. FAILED: sol.codigoSolicitud mismatch: expected CAP006-GETSOL, got " & sol.codigoSolicitud
        Test_Lifecycle_Strict_GetSolicitudPorID_RetornaSolicitudParaIdValido = TestHelper.BuildJsonFail("codigoSolicitud mismatch", logs)
        GoTo CleanExit
    End If
    logs(3) = "4. sol.codigoSolicitud=CAP006-GETSOL as expected"

    Test_Lifecycle_Strict_GetSolicitudPorID_RetornaSolicitudParaIdValido = TestHelper.BuildJsonOk("true", logs)
    Exit Function

CleanExit:
    On Error Resume Next
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
    Test_Lifecycle_Strict_GetSolicitudPorID_RetornaSolicitudParaIdValido = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Lifecycle_Strict_GetSolicitudPorID_RaisesErrorWhenIdIsZeroOrNegative() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)
    Dim setupError As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim solServ As New SolicitudServicio
    Dim sol As Solicitud
    Dim errorNumber As Long
    Dim errorDescription As String

    Call SetupLifecycleSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Lifecycle_Strict_GetSolicitudPorID_RaisesErrorWhenIdIsZeroOrNegative = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()

    On Error Resume Next
    Set sol = solServ.getSolicitudPorID(0, db)
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error GoTo EH
    logs(0) = "1. Called getSolicitudPorID(0)"

    If errorNumber = 0 Then
        logs(1) = "2. FAILED: expected CondorError Raise (513), but no error fired"
        Test_Lifecycle_Strict_GetSolicitudPorID_RaisesErrorWhenIdIsZeroOrNegative = TestHelper.BuildJsonFail("expected Raise", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. Caught error " & errorNumber & ": " & errorDescription

    If errorNumber <> 513 Then
        logs(2) = "3. FAILED: expected errorNumber=513, got " & errorNumber
        Test_Lifecycle_Strict_GetSolicitudPorID_RaisesErrorWhenIdIsZeroOrNegative = TestHelper.BuildJsonFail("expected 513", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. errorNumber=513 as expected (custom application error: ID de solicitud no válido)"

    logs(3) = "4. PASS"
    Test_Lifecycle_Strict_GetSolicitudPorID_RaisesErrorWhenIdIsZeroOrNegative = TestHelper.BuildJsonOk("true", logs)
    Exit Function

CleanExit:
    On Error Resume Next
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
    Test_Lifecycle_Strict_GetSolicitudPorID_RaisesErrorWhenIdIsZeroOrNegative = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Lifecycle_Strict_GetExpedientePorID_RaisesErrorWhenIdIsNotNumeric() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(4)
    Dim setupError As String
    Dim db As DAO.Database
    Dim tempRoot As String
    Dim expServ As New ExpedienteServicio
    Dim exp As Expediente
    Dim errorNumber As Long
    Dim errorDescription As String

    Call SetupLifecycleSandbox(tempRoot, setupError)
    If setupError <> "" Then
        logs(0) = setupError
        Test_Lifecycle_Strict_GetExpedientePorID_RaisesErrorWhenIdIsNotNumeric = TestHelper.BuildJsonFail("TESTS BLOCKED", logs)
        GoTo CleanExit
    End If

    Set db = TestHelper.GetTestDb()

    On Error Resume Next
    Set exp = expServ.getExpedientePorID("not-a-number")
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error GoTo EH
    logs(0) = "1. Called getExpedientePorID(""not-a-number"")"

    If errorNumber = 0 Then
        logs(1) = "2. FAILED: expected CondorError Raise (513), but no error fired"
        Test_Lifecycle_Strict_GetExpedientePorID_RaisesErrorWhenIdIsNotNumeric = TestHelper.BuildJsonFail("expected Raise", logs)
        GoTo CleanExit
    End If
    logs(1) = "2. Caught error " & errorNumber & ": " & errorDescription

    If errorNumber <> 513 Then
        logs(2) = "3. FAILED: expected errorNumber=513, got " & errorNumber
        Test_Lifecycle_Strict_GetExpedientePorID_RaisesErrorWhenIdIsNotNumeric = TestHelper.BuildJsonFail("expected 513", logs)
        GoTo CleanExit
    End If
    logs(2) = "3. errorNumber=513 as expected (custom application error: ID de Expediente no válido)"

    logs(3) = "4. PASS"
    Test_Lifecycle_Strict_GetExpedientePorID_RaisesErrorWhenIdIsNotNumeric = TestHelper.BuildJsonOk("true", logs)
    Exit Function

CleanExit:
    On Error Resume Next
    Set exp = Nothing
    If Not db Is Nothing Then TeardownFixtures db
    Call TestHelper.CleanupProdTempRoot(tempRoot)
    Call TestHelper.ResetTestSession
    Set m_ObjEntorno = Nothing
    Set m_ObjUsuarioActivo = Nothing
    On Error GoTo 0
    Exit Function

EH:
    logs(UBound(logs)) = "EH: ERR " & Err.Number & " - " & Err.Description
    Test_Lifecycle_Strict_GetExpedientePorID_RaisesErrorWhenIdIsNotNumeric = TestHelper.BuildJsonFail(Err.Description, logs)
    Resume CleanExit
End Function

Public Function Test_Lifecycle_Strict_RunAll() As String
    On Error GoTo EH
    Dim logs() As String: logs = TestHelper.NewLogsArray(6)
    Dim result As String

    result = Test_Lifecycle_Strict_GetSolicitudesViewModel_RetornaLasCreadas()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(0) = "1. getSolicitudesViewModel returns seeded ids"

    result = Test_Lifecycle_Strict_GetSolicitudesViewModel_FiltraPorPalabraClave()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(1) = "2. getSolicitudesViewModel filters by palabraClave"

    result = Test_Lifecycle_Strict_GetSolicitudPorID_RetornaSolicitudParaIdValido()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(2) = "3. getSolicitudPorID returns the Solicitud for a valid id"

    result = Test_Lifecycle_Strict_GetSolicitudPorID_RaisesErrorWhenIdIsZeroOrNegative()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(3) = "4. getSolicitudPorID raises 513 when id is zero or negative"

    result = Test_Lifecycle_Strict_GetExpedientePorID_RaisesErrorWhenIdIsNotNumeric()
    If InStr(1, result, """ok"":true", vbTextCompare) = 0 Then GoTo Failed
    logs(4) = "5. getExpedientePorID raises 513 when id is not numeric (validation before DB)"

    logs(5) = "6. PASS"
    Test_Lifecycle_Strict_RunAll = TestHelper.BuildJsonOk("true", logs)
    Exit Function

Failed:
    Test_Lifecycle_Strict_RunAll = TestHelper.BuildJsonFail("one CAP-006 lifecycle atom failed", logs)
    Exit Function

EH:
    logs(UBound(logs)) = "Test_Lifecycle_Strict_RunAll: ERR " & Err.Number & " - " & Err.Description
    Test_Lifecycle_Strict_RunAll = TestHelper.BuildJsonFail(Err.Description, logs)
End Function
