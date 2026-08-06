Attribute VB_Name = "Test_CacheConsistencyAudit"
Option Compare Database
Option Explicit

Private Const TEST_USER_DELETE_REMAINING As Long = 901101
Private Const TEST_USER_DELETE_LAST As Long = 901102
Private Const TEST_USER_EARLIER_MIN As Long = 901103
Private Const TEST_USER_NULL_MIN As Long = 901104
Private Const TEST_SICA_DELETE As String = "SICA901201"
Private Const TEST_SICA_MISSING_LOCAL As String = "SICA901202"
Private Const TEST_SICA_TO_HISTORICAL As String = "SICA901203"
Private Const TEST_SICA_TO_CURRENT As String = "SICA901204"
Private Const TEST_SICA_CURRENT_HPS As Long = 901301
Private Const TEST_SICA_HISTORICAL_HPS As Long = 901302

Public Function Test_CCA_HpsDeleteWithRemaining_AlignsCaches() As String
    Dim logs As Collection
    Dim sourceDb As DAO.Database
    Dim localDb As DAO.Database
    Dim errMsg As String
    Dim beforeList As Long
    Dim beforeIndicators As Long
    Dim afterList As Long
    Dim afterIndicators As Long
    Set logs = New Collection
    On Error GoTo EH

    Set localDb = CurrentDb()
    Set sourceDb = OpenBackendSourceDb()
    SeedHpsFixture sourceDb, localDb, TEST_USER_DELETE_REMAINING, True, #1/10/2024#, #2/20/2024#, #12/31/2023#
    beforeList = CountRows(localDb, "TbDatosLocal", "ID=" & TEST_USER_DELETE_REMAINING)
    beforeIndicators = CountRows(localDb, "TbDatosLocalParaIndicadores", "ID=" & TEST_USER_DELETE_REMAINING)

    sourceDb.Execute "DELETE * FROM TbHPS WHERE IDUsuario=" & TEST_USER_DELETE_REMAINING & " AND TipoHPS='Nacional'", dbFailOnError
    Call RefreshHpsUserCacheAfterMutation(CStr(TEST_USER_DELETE_REMAINING), sourceDb, localDb, errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    afterList = CountRows(localDb, "TbDatosLocal", "ID=" & TEST_USER_DELETE_REMAINING)
    afterIndicators = CountRows(localDb, "TbDatosLocalParaIndicadores", "ID=" & TEST_USER_DELETE_REMAINING)
    If afterList <> 1 Or afterIndicators <> 1 Then
        Test_CCA_HpsDeleteWithRemaining_AlignsCaches = JsonFail("Expected one aligned cache row after deleting one HPS while another remains.", logs)
        GoTo CleanUp
    End If
    If Not CacheDateEquals(localDb, "TbDatosLocal", TEST_USER_DELETE_REMAINING, #2/20/2024#) Then
        Test_CCA_HpsDeleteWithRemaining_AlignsCaches = JsonFail("TbDatosLocal did not expose remaining HPS min date.", logs)
        GoTo CleanUp
    End If
    If Not CacheDateEquals(localDb, "TbDatosLocalParaIndicadores", TEST_USER_DELETE_REMAINING, #2/20/2024#) Then
        Test_CCA_HpsDeleteWithRemaining_AlignsCaches = JsonFail("TbDatosLocalParaIndicadores did not expose remaining HPS min date.", logs)
        GoTo CleanUp
    End If

    logs.Add "Deleting one HPS retained one list row and one indicator row."
    Test_CCA_HpsDeleteWithRemaining_AlignsCaches = JsonOkPayload("hps-delete-with-remaining", CardinalityPayload(beforeList, beforeIndicators, afterList, afterIndicators), logs)
CleanUp:
    TeardownHpsFixture sourceDb, localDb, TEST_USER_DELETE_REMAINING
    CloseSourceDb sourceDb
    Exit Function
EH:
    Test_CCA_HpsDeleteWithRemaining_AlignsCaches = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownHpsFixture sourceDb, localDb, TEST_USER_DELETE_REMAINING
    CloseSourceDb sourceDb
End Function

Public Function Test_CCA_HpsDeleteLast_SourceEquivalentCaches() As String
    Dim logs As Collection
    Dim sourceDb As DAO.Database
    Dim localDb As DAO.Database
    Dim errMsg As String
    Dim beforeList As Long
    Dim beforeIndicators As Long
    Dim afterList As Long
    Dim afterIndicators As Long
    Set logs = New Collection
    On Error GoTo EH

    Set localDb = CurrentDb()
    Set sourceDb = OpenBackendSourceDb()
    SeedHpsFixture sourceDb, localDb, TEST_USER_DELETE_LAST, False, #3/15/2024#, Empty, #3/15/2024#
    beforeList = CountRows(localDb, "TbDatosLocal", "ID=" & TEST_USER_DELETE_LAST)
    beforeIndicators = CountRows(localDb, "TbDatosLocalParaIndicadores", "ID=" & TEST_USER_DELETE_LAST)

    sourceDb.Execute "DELETE * FROM TbHPS WHERE IDUsuario=" & TEST_USER_DELETE_LAST, dbFailOnError
    Call RefreshHpsUserCacheAfterMutation(CStr(TEST_USER_DELETE_LAST), sourceDb, localDb, errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    afterList = CountRows(localDb, "TbDatosLocal", "ID=" & TEST_USER_DELETE_LAST)
    afterIndicators = CountRows(localDb, "TbDatosLocalParaIndicadores", "ID=" & TEST_USER_DELETE_LAST)
    If afterList <> 0 Or afterIndicators <> 0 Then
        Test_CCA_HpsDeleteLast_SourceEquivalentCaches = JsonFail("Last HPS delete should leave no user cache rows.", logs)
        GoTo CleanUp
    End If

    logs.Add "Deleting the last HPS removed list and indicator cache rows."
    Test_CCA_HpsDeleteLast_SourceEquivalentCaches = JsonOkPayload("hps-delete-last", CardinalityPayload(beforeList, beforeIndicators, afterList, afterIndicators), logs)
CleanUp:
    TeardownHpsFixture sourceDb, localDb, TEST_USER_DELETE_LAST
    CloseSourceDb sourceDb
    Exit Function
EH:
    Test_CCA_HpsDeleteLast_SourceEquivalentCaches = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownHpsFixture sourceDb, localDb, TEST_USER_DELETE_LAST
    CloseSourceDb sourceDb
End Function

Public Function Test_CCA_HpsEarlierMinDate_AlignsCaches() As String
    Dim logs As Collection
    Dim sourceDb As DAO.Database
    Dim localDb As DAO.Database
    Dim errMsg As String
    Set logs = New Collection
    On Error GoTo EH

    Set localDb = CurrentDb()
    Set sourceDb = OpenBackendSourceDb()
    SeedHpsFixture sourceDb, localDb, TEST_USER_EARLIER_MIN, False, #5/10/2024#, Empty, #5/10/2024#
    sourceDb.Execute "INSERT INTO TbHPS (IDUsuario, TipoHPS, F_Concesion) VALUES (" & TEST_USER_EARLIER_MIN & ", 'UE', #1/5/2024#)", dbFailOnError
    Call RefreshHpsUserCacheAfterMutation(CStr(TEST_USER_EARLIER_MIN), sourceDb, localDb, errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    If Not CacheDateEquals(localDb, "TbDatosLocal", TEST_USER_EARLIER_MIN, #1/5/2024#) Then
        Test_CCA_HpsEarlierMinDate_AlignsCaches = JsonFail("TbDatosLocal kept the previous minimum date.", logs)
        GoTo CleanUp
    End If
    If Not CacheDateEquals(localDb, "TbDatosLocalParaIndicadores", TEST_USER_EARLIER_MIN, #1/5/2024#) Then
        Test_CCA_HpsEarlierMinDate_AlignsCaches = JsonFail("TbDatosLocalParaIndicadores kept the previous minimum date.", logs)
        GoTo CleanUp
    End If

    logs.Add "Earlier source concession date propagated to both HPS cache tables."
    Test_CCA_HpsEarlierMinDate_AlignsCaches = JsonOkPayload("hps-earlier-min-date", CardinalityPayload(1, 1, CountRows(localDb, "TbDatosLocal", "ID=" & TEST_USER_EARLIER_MIN), CountRows(localDb, "TbDatosLocalParaIndicadores", "ID=" & TEST_USER_EARLIER_MIN)), logs)
CleanUp:
    TeardownHpsFixture sourceDb, localDb, TEST_USER_EARLIER_MIN
    CloseSourceDb sourceDb
    Exit Function
EH:
    Test_CCA_HpsEarlierMinDate_AlignsCaches = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownHpsFixture sourceDb, localDb, TEST_USER_EARLIER_MIN
    CloseSourceDb sourceDb
End Function

Public Function Test_CCA_HpsNullMinDate_AlignsCaches() As String
    Dim logs As Collection
    Dim sourceDb As DAO.Database
    Dim localDb As DAO.Database
    Dim errMsg As String
    Set logs = New Collection
    On Error GoTo EH

    Set localDb = CurrentDb()
    Set sourceDb = OpenBackendSourceDb()
    SeedHpsFixture sourceDb, localDb, TEST_USER_NULL_MIN, False, Empty, Empty, #4/4/2024#
    Call RefreshHpsUserCacheAfterMutation(CStr(TEST_USER_NULL_MIN), sourceDb, localDb, errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    If Not CacheDateIsNull(localDb, "TbDatosLocal", TEST_USER_NULL_MIN) Then
        Test_CCA_HpsNullMinDate_AlignsCaches = JsonFail("TbDatosLocal did not clear empty source min date.", logs)
        GoTo CleanUp
    End If
    If Not CacheDateIsNull(localDb, "TbDatosLocalParaIndicadores", TEST_USER_NULL_MIN) Then
        Test_CCA_HpsNullMinDate_AlignsCaches = JsonFail("TbDatosLocalParaIndicadores did not clear empty source min date.", logs)
        GoTo CleanUp
    End If

    logs.Add "Null/no-date source state cleared both HPS cache dates."
    Test_CCA_HpsNullMinDate_AlignsCaches = JsonOkPayload("hps-null-min-date", CardinalityPayload(1, 1, CountRows(localDb, "TbDatosLocal", "ID=" & TEST_USER_NULL_MIN), CountRows(localDb, "TbDatosLocalParaIndicadores", "ID=" & TEST_USER_NULL_MIN)), logs)
CleanUp:
    TeardownHpsFixture sourceDb, localDb, TEST_USER_NULL_MIN
    CloseSourceDb sourceDb
    Exit Function
EH:
    Test_CCA_HpsNullMinDate_AlignsCaches = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownHpsFixture sourceDb, localDb, TEST_USER_NULL_MIN
    CloseSourceDb sourceDb
End Function

Public Function Test_CCA_SicaDelete_RemovesLocalAndIndicatorCaches() As String
    Dim logs As Collection
    Dim sourceDb As DAO.Database
    Dim localDb As DAO.Database
    Dim errMsg As String
    Dim beforeList As Long
    Dim beforeIndicators As Long
    Dim afterList As Long
    Dim afterIndicators As Long
    Set logs = New Collection
    On Error GoTo EH

    Set localDb = CurrentDb()
    Set sourceDb = OpenBackendSourceDb()
    SeedSicaFixture sourceDb, localDb, TEST_SICA_DELETE, TEST_SICA_CURRENT_HPS, 0, "Stale Delete"
    beforeList = CountRows(localDb, "TbUsuariosSICALocal", "ID='" & TEST_SICA_DELETE & "'")
    beforeIndicators = CountRows(localDb, "TbUsuariosSICALocalParaIndicadores", "ID='" & TEST_SICA_DELETE & "'")

    sourceDb.Execute "DELETE * FROM TbUsuariosSICA WHERE ID='" & TEST_SICA_DELETE & "'", dbFailOnError
    Call DeleteSicaLocalCaches(TEST_SICA_DELETE, sourceDb, localDb, errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    afterList = CountRows(localDb, "TbUsuariosSICALocal", "ID='" & TEST_SICA_DELETE & "'")
    afterIndicators = CountRows(localDb, "TbUsuariosSICALocalParaIndicadores", "ID='" & TEST_SICA_DELETE & "'")
    If afterList <> 0 Or afterIndicators <> 0 Then
        Test_CCA_SicaDelete_RemovesLocalAndIndicatorCaches = JsonFail("SICA delete left stale local or indicator cache rows.", logs)
        GoTo CleanUp
    End If

    logs.Add "Deleting a SICA source row cleared both SICA cache tables."
    Test_CCA_SicaDelete_RemovesLocalAndIndicatorCaches = JsonOkPayload("sica-delete", CardinalityPayload(beforeList, beforeIndicators, afterList, afterIndicators), logs)
CleanUp:
    TeardownSicaFixture sourceDb, localDb, TEST_SICA_DELETE, TEST_SICA_CURRENT_HPS, 0
    CloseSourceDb sourceDb
    Exit Function
EH:
    Test_CCA_SicaDelete_RemovesLocalAndIndicatorCaches = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownSicaFixture sourceDb, localDb, TEST_SICA_DELETE, TEST_SICA_CURRENT_HPS, 0
    CloseSourceDb sourceDb
End Function

Public Function Test_CCA_SicaDelete_MissingLocalClearsIndicatorSafely() As String
    Dim logs As Collection
    Dim sourceDb As DAO.Database
    Dim localDb As DAO.Database
    Dim errMsg As String
    Dim beforeList As Long
    Dim beforeIndicators As Long
    Dim afterList As Long
    Dim afterIndicators As Long
    Set logs = New Collection
    On Error GoTo EH

    Set localDb = CurrentDb()
    Set sourceDb = OpenBackendSourceDb()
    TeardownSicaFixture sourceDb, localDb, TEST_SICA_MISSING_LOCAL, TEST_SICA_CURRENT_HPS, 0
    InsertSicaCacheRow localDb, "TbUsuariosSICALocalParaIndicadores", TEST_SICA_MISSING_LOCAL, TEST_SICA_CURRENT_HPS, 0, "Indicator Only"
    beforeList = CountRows(localDb, "TbUsuariosSICALocal", "ID='" & TEST_SICA_MISSING_LOCAL & "'")
    beforeIndicators = CountRows(localDb, "TbUsuariosSICALocalParaIndicadores", "ID='" & TEST_SICA_MISSING_LOCAL & "'")

    Call DeleteSicaLocalCaches(TEST_SICA_MISSING_LOCAL, sourceDb, localDb, errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    afterList = CountRows(localDb, "TbUsuariosSICALocal", "ID='" & TEST_SICA_MISSING_LOCAL & "'")
    afterIndicators = CountRows(localDb, "TbUsuariosSICALocalParaIndicadores", "ID='" & TEST_SICA_MISSING_LOCAL & "'")
    If beforeList <> 0 Or beforeIndicators <> 1 Or afterList <> 0 Or afterIndicators <> 0 Then
        Test_CCA_SicaDelete_MissingLocalClearsIndicatorSafely = JsonFail("Missing local row scenario did not clear stale SICA indicator cache safely.", logs)
        GoTo CleanUp
    End If

    logs.Add "Missing local SICA row did not block stale indicator cleanup."
    Test_CCA_SicaDelete_MissingLocalClearsIndicatorSafely = JsonOkPayload("sica-delete-missing-local", CardinalityPayload(beforeList, beforeIndicators, afterList, afterIndicators), logs)
CleanUp:
    TeardownSicaFixture sourceDb, localDb, TEST_SICA_MISSING_LOCAL, TEST_SICA_CURRENT_HPS, 0
    CloseSourceDb sourceDb
    Exit Function
EH:
    Test_CCA_SicaDelete_MissingLocalClearsIndicatorSafely = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownSicaFixture sourceDb, localDb, TEST_SICA_MISSING_LOCAL, TEST_SICA_CURRENT_HPS, 0
    CloseSourceDb sourceDb
End Function

Public Function Test_CCA_SicaCurrentToHistorical_RefreshesCaches() As String
    Dim logs As Collection
    Dim sourceDb As DAO.Database
    Dim localDb As DAO.Database
    Dim errMsg As String
    Set logs = New Collection
    On Error GoTo EH

    Set localDb = CurrentDb()
    Set sourceDb = OpenBackendSourceDb()
    SeedSicaFixture sourceDb, localDb, TEST_SICA_TO_HISTORICAL, TEST_SICA_CURRENT_HPS, 0, "Current Link"

    sourceDb.Execute "UPDATE TbUsuariosSICA SET IDHPS=Null, IDHPSHistorico=" & TEST_SICA_HISTORICAL_HPS & " WHERE ID='" & TEST_SICA_TO_HISTORICAL & "'", dbFailOnError
    Call RefreshSicaLocalCaches(TEST_SICA_TO_HISTORICAL, sourceDb, localDb, errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    If Not SicaCacheLinkEquals(localDb, "TbUsuariosSICALocal", TEST_SICA_TO_HISTORICAL, 0, TEST_SICA_HISTORICAL_HPS) Then
        Test_CCA_SicaCurrentToHistorical_RefreshesCaches = JsonFail("TbUsuariosSICALocal did not expose the historical SICA link.", logs)
        GoTo CleanUp
    End If
    If Not SicaCacheLinkEquals(localDb, "TbUsuariosSICALocalParaIndicadores", TEST_SICA_TO_HISTORICAL, 0, TEST_SICA_HISTORICAL_HPS) Then
        Test_CCA_SicaCurrentToHistorical_RefreshesCaches = JsonFail("TbUsuariosSICALocalParaIndicadores did not expose the historical SICA link.", logs)
        GoTo CleanUp
    End If

    logs.Add "SICA current-to-historical link refresh aligned local and indicator caches."
    Test_CCA_SicaCurrentToHistorical_RefreshesCaches = JsonOkPayload("sica-current-to-historical", SicaLinkPayload(localDb, TEST_SICA_TO_HISTORICAL), logs)
CleanUp:
    TeardownSicaFixture sourceDb, localDb, TEST_SICA_TO_HISTORICAL, TEST_SICA_CURRENT_HPS, TEST_SICA_HISTORICAL_HPS
    CloseSourceDb sourceDb
    Exit Function
EH:
    Test_CCA_SicaCurrentToHistorical_RefreshesCaches = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownSicaFixture sourceDb, localDb, TEST_SICA_TO_HISTORICAL, TEST_SICA_CURRENT_HPS, TEST_SICA_HISTORICAL_HPS
    CloseSourceDb sourceDb
End Function

Public Function Test_CCA_SicaHistoricalToCurrent_RefreshesCaches() As String
    Dim logs As Collection
    Dim sourceDb As DAO.Database
    Dim localDb As DAO.Database
    Dim errMsg As String
    Set logs = New Collection
    On Error GoTo EH

    Set localDb = CurrentDb()
    Set sourceDb = OpenBackendSourceDb()
    SeedSicaFixture sourceDb, localDb, TEST_SICA_TO_CURRENT, 0, TEST_SICA_HISTORICAL_HPS, "Historical Link"

    sourceDb.Execute "UPDATE TbUsuariosSICA SET IDHPS=" & TEST_SICA_CURRENT_HPS & ", IDHPSHistorico=Null WHERE ID='" & TEST_SICA_TO_CURRENT & "'", dbFailOnError
    Call RefreshSicaLocalCaches(TEST_SICA_TO_CURRENT, sourceDb, localDb, errMsg)
    If errMsg <> "" Then Err.Raise 1000, , errMsg

    If Not SicaCacheLinkEquals(localDb, "TbUsuariosSICALocal", TEST_SICA_TO_CURRENT, TEST_SICA_CURRENT_HPS, 0) Then
        Test_CCA_SicaHistoricalToCurrent_RefreshesCaches = JsonFail("TbUsuariosSICALocal did not expose the current SICA link.", logs)
        GoTo CleanUp
    End If
    If Not SicaCacheLinkEquals(localDb, "TbUsuariosSICALocalParaIndicadores", TEST_SICA_TO_CURRENT, TEST_SICA_CURRENT_HPS, 0) Then
        Test_CCA_SicaHistoricalToCurrent_RefreshesCaches = JsonFail("TbUsuariosSICALocalParaIndicadores did not expose the current SICA link.", logs)
        GoTo CleanUp
    End If

    logs.Add "SICA historical-to-current link refresh aligned local and indicator caches."
    Test_CCA_SicaHistoricalToCurrent_RefreshesCaches = JsonOkPayload("sica-historical-to-current", SicaLinkPayload(localDb, TEST_SICA_TO_CURRENT), logs)
CleanUp:
    TeardownSicaFixture sourceDb, localDb, TEST_SICA_TO_CURRENT, TEST_SICA_CURRENT_HPS, TEST_SICA_HISTORICAL_HPS
    CloseSourceDb sourceDb
    Exit Function
EH:
    Test_CCA_SicaHistoricalToCurrent_RefreshesCaches = JsonFail("Unexpected error: " & Err.Description, logs)
    On Error Resume Next
    TeardownSicaFixture sourceDb, localDb, TEST_SICA_TO_CURRENT, TEST_SICA_CURRENT_HPS, TEST_SICA_HISTORICAL_HPS
    CloseSourceDb sourceDb
End Function

Private Sub SeedHpsFixture(ByRef sourceDb As DAO.Database, ByRef localDb As DAO.Database, ByVal userId As Long, ByVal includeSecondHps As Boolean, ByVal firstDate As Variant, ByVal secondDate As Variant, ByVal staleCacheDate As Variant)
    TeardownHpsFixture sourceDb, localDb, userId
    sourceDb.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & userId & ", 'DNI" & userId & "', 'Cache', 'Audit')", dbFailOnError
    InsertHps sourceDb, userId, "Nacional", firstDate
    If includeSecondHps Then InsertHps sourceDb, userId, "OTAN", secondDate
    InsertCacheRows localDb, userId, staleCacheDate
End Sub

Private Sub InsertHps(ByRef db As DAO.Database, ByVal userId As Long, ByVal tipoHps As String, ByVal concessionDate As Variant)
    Dim sql As String
    If IsDate(concessionDate) Then
        sql = "INSERT INTO TbHPS (IDUsuario, TipoHPS, F_Concesion) VALUES (" & userId & ", '" & tipoHps & "', #" & Format$(CDate(concessionDate), "mm/dd/yyyy") & "#)"
    Else
        sql = "INSERT INTO TbHPS (IDUsuario, TipoHPS) VALUES (" & userId & ", '" & tipoHps & "')"
    End If
    db.Execute sql, dbFailOnError
End Sub

Private Sub InsertCacheRows(ByRef db As DAO.Database, ByVal userId As Long, ByVal cacheDate As Variant)
    InsertCacheRow db, "TbDatosLocal", userId, cacheDate
    InsertCacheRow db, "TbDatosLocalParaIndicadores", userId, cacheDate
End Sub

Private Sub InsertCacheRow(ByRef db As DAO.Database, ByVal tableName As String, ByVal userId As Long, ByVal cacheDate As Variant)
    Dim sql As String
    If IsDate(cacheDate) Then
        sql = "INSERT INTO " & tableName & " (ID, DNI, Nombre, Apellido_1, FechaHPSConcesionMinima) VALUES (" & userId & ", 'DNI" & userId & "', 'Stale', 'Cache', #" & Format$(CDate(cacheDate), "mm/dd/yyyy") & "#)"
    Else
        sql = "INSERT INTO " & tableName & " (ID, DNI, Nombre, Apellido_1) VALUES (" & userId & ", 'DNI" & userId & "', 'Stale', 'Cache')"
    End If
    db.Execute sql, dbFailOnError
End Sub

Private Sub TeardownHpsFixture(ByRef sourceDb As DAO.Database, ByRef localDb As DAO.Database, ByVal userId As Long)
    On Error Resume Next
    If Not sourceDb Is Nothing Then
        sourceDb.Execute "DELETE * FROM TbHPS WHERE IDUsuario=" & userId, dbFailOnError
        sourceDb.Execute "DELETE * FROM TbUsuariosEntidades WHERE ID=" & userId, dbFailOnError
        sourceDb.Execute "DELETE * FROM TbUsuarios WHERE ID=" & userId, dbFailOnError
    End If
    If Not localDb Is Nothing Then
        localDb.Execute "DELETE * FROM TbDatosLocal WHERE ID=" & userId, dbFailOnError
        localDb.Execute "DELETE * FROM TbDatosLocalParaIndicadores WHERE ID=" & userId, dbFailOnError
    End If
End Sub

Private Sub SeedSicaFixture(ByRef sourceDb As DAO.Database, ByRef localDb As DAO.Database, ByVal sicaId As String, ByVal currentHpsId As Long, ByVal historicalHpsId As Long, ByVal displayName As String)
    TeardownSicaFixture sourceDb, localDb, sicaId, currentHpsId, historicalHpsId
    If currentHpsId > 0 Then InsertUsuario sourceDb, currentHpsId, displayName & " Current"
    If historicalHpsId > 0 Then InsertUsuarioHistorico sourceDb, historicalHpsId, displayName & " Historical"
    InsertSicaSourceRow sourceDb, sicaId, currentHpsId, historicalHpsId, displayName
    InsertSicaCacheRow localDb, "TbUsuariosSICALocal", sicaId, currentHpsId, historicalHpsId, "Stale " & displayName
    InsertSicaCacheRow localDb, "TbUsuariosSICALocalParaIndicadores", sicaId, currentHpsId, historicalHpsId, "Stale " & displayName
End Sub

Private Sub InsertUsuario(ByRef db As DAO.Database, ByVal userId As Long, ByVal displayName As String)
    db.Execute "INSERT INTO TbUsuarios (ID, DNI, Nombre, Apellido_1) VALUES (" & userId & ", 'DNI" & userId & "', '" & SqlText(displayName) & "', 'Cache')", dbFailOnError
End Sub

Private Sub InsertUsuarioHistorico(ByRef db As DAO.Database, ByVal userId As Long, ByVal displayName As String)
    db.Execute "INSERT INTO TbUsuariosHistoricos (ID, DNI, Nombre, Apellido_1) VALUES (" & userId & ", 'DNI" & userId & "', '" & SqlText(displayName) & "', 'Cache')", dbFailOnError
End Sub

Private Sub InsertSicaSourceRow(ByRef db As DAO.Database, ByVal sicaId As String, ByVal currentHpsId As Long, ByVal historicalHpsId As Long, ByVal displayName As String)
    Dim sql As String
    sql = "INSERT INTO TbUsuariosSICA (ID, IDHPS, IDHPSHistorico, Nombre, Apellido_1, TramitacionExterna, UsuarioSICATSOL, FechaCreacion, FechaModificacion, Correo_e) VALUES (" & _
          "'" & SqlText(sicaId) & "', " & NullableLong(currentHpsId) & ", " & NullableLong(historicalHpsId) & ", '" & SqlText(displayName) & "', 'Cache', 'No', 'USR" & Right$(sicaId, 4) & "', #1/1/2024#, #1/1/2024#, '" & LCase$(sicaId) & "@example.test')"
    db.Execute sql, dbFailOnError
End Sub

Private Sub InsertSicaCacheRow(ByRef db As DAO.Database, ByVal tableName As String, ByVal sicaId As String, ByVal currentHpsId As Long, ByVal historicalHpsId As Long, ByVal displayName As String)
    Dim sql As String
    sql = "INSERT INTO " & tableName & " (ID, IDHPS, IDHPSHistorico, Nombre, Apellido_1, TramitacionExterna, UsuarioSICATSOL, FechaCreacion, FechaModificacion, Correo_e) VALUES (" & _
          "'" & SqlText(sicaId) & "', " & NullableLong(currentHpsId) & ", " & NullableLong(historicalHpsId) & ", '" & SqlText(displayName) & "', 'Cache', 'No', 'USR" & Right$(sicaId, 4) & "', #1/1/2024#, #1/1/2024#, '" & LCase$(sicaId) & "@example.test')"
    db.Execute sql, dbFailOnError
End Sub

Private Sub TeardownSicaFixture(ByRef sourceDb As DAO.Database, ByRef localDb As DAO.Database, ByVal sicaId As String, ByVal currentHpsId As Long, ByVal historicalHpsId As Long)
    On Error Resume Next
    If Not sourceDb Is Nothing Then
        sourceDb.Execute "DELETE * FROM TbUsuariosSICA WHERE ID='" & SqlText(sicaId) & "'", dbFailOnError
        If currentHpsId > 0 Then sourceDb.Execute "DELETE * FROM TbUsuarios WHERE ID=" & currentHpsId, dbFailOnError
        If historicalHpsId > 0 Then sourceDb.Execute "DELETE * FROM TbUsuariosHistoricos WHERE ID=" & historicalHpsId, dbFailOnError
    End If
    If Not localDb Is Nothing Then
        localDb.Execute "DELETE * FROM TbUsuariosSICALocal WHERE ID='" & SqlText(sicaId) & "'", dbFailOnError
        localDb.Execute "DELETE * FROM TbUsuariosSICALocalParaIndicadores WHERE ID='" & SqlText(sicaId) & "'", dbFailOnError
    End If
End Sub

Private Function SicaCacheLinkEquals(ByRef db As DAO.Database, ByVal tableName As String, ByVal sicaId As String, ByVal expectedCurrentHpsId As Long, ByVal expectedHistoricalHpsId As Long) As Boolean
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT IDHPS, IDHPSHistorico FROM " & tableName & " WHERE ID='" & SqlText(sicaId) & "'", dbOpenSnapshot)
    If Not rs.EOF Then
        SicaCacheLinkEquals = (LongOrZero(rs!IDHPS) = expectedCurrentHpsId And LongOrZero(rs!IDHPSHistorico) = expectedHistoricalHpsId)
    End If
    rs.Close
End Function

Private Function SicaLinkPayload(ByRef db As DAO.Database, ByVal sicaId As String) As String
    Dim rsLocal As DAO.Recordset
    Dim rsIndicators As DAO.Recordset
    Set rsLocal = db.OpenRecordset("SELECT IDHPS, IDHPSHistorico FROM TbUsuariosSICALocal WHERE ID='" & SqlText(sicaId) & "'", dbOpenSnapshot)
    Set rsIndicators = db.OpenRecordset("SELECT IDHPS, IDHPSHistorico FROM TbUsuariosSICALocalParaIndicadores WHERE ID='" & SqlText(sicaId) & "'", dbOpenSnapshot)
    SicaLinkPayload = "{""localCurrent"":" & LongOrZero(rsLocal!IDHPS) & ",""localHistorical"":" & LongOrZero(rsLocal!IDHPSHistorico) & ",""indicatorCurrent"":" & LongOrZero(rsIndicators!IDHPS) & ",""indicatorHistorical"":" & LongOrZero(rsIndicators!IDHPSHistorico) & "}"
    rsLocal.Close
    rsIndicators.Close
End Function

Private Function LongOrZero(ByVal value As Variant) As Long
    If IsNull(value) Or value = "" Then
        LongOrZero = 0
    Else
        LongOrZero = CLng(value)
    End If
End Function

Private Function NullableLong(ByVal value As Long) As String
    If value > 0 Then
        NullableLong = CStr(value)
    Else
        NullableLong = "Null"
    End If
End Function

Private Function SqlText(ByVal value As String) As String
    SqlText = Replace(value, "'", "''")
End Function

Private Function OpenBackendSourceDb() As DAO.Database
    Set OpenBackendSourceDb = DBEngine.Workspaces(0).OpenDatabase(CurrentProject.path & "\HPST.accdb", False, False, BackendConnectString())
End Function

Private Function BackendConnectString() As String
    Dim password As String
    password = Environ$("DYSFLOW_BACKEND_PASSWORD")
    If Len(password) = 0 Then password = Environ$("HPS_BACKEND_PASSWORD")
    If Len(password) = 0 Then password = Environ$("ACCESS_VBA_PASSWORD")
    If Len(password) = 0 Then Err.Raise 1000, , "BackendConnectString: DYSFLOW_BACKEND_PASSWORD, HPS_BACKEND_PASSWORD, or ACCESS_VBA_PASSWORD is required for cache consistency tests."
    BackendConnectString = "MS Access;PWD=" & password
End Function

Private Sub CloseSourceDb(ByRef sourceDb As DAO.Database)
    On Error Resume Next
    If Not sourceDb Is Nothing Then sourceDb.Close
    Set sourceDb = Nothing
End Sub

Private Function CountRows(ByRef db As DAO.Database, ByVal tableName As String, ByVal whereClause As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM " & tableName & " WHERE " & whereClause, dbOpenSnapshot)
    CountRows = CLng(rs!C)
    rs.Close
End Function

Private Function CacheDateEquals(ByRef db As DAO.Database, ByVal tableName As String, ByVal userId As Long, ByVal expectedDate As Date) As Boolean
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT FechaHPSConcesionMinima FROM " & tableName & " WHERE ID=" & userId, dbOpenSnapshot)
    If Not rs.EOF Then
        If IsDate(rs!FechaHPSConcesionMinima) Then CacheDateEquals = (DateValue(rs!FechaHPSConcesionMinima) = DateValue(expectedDate))
    End If
    rs.Close
End Function

Private Function CacheDateIsNull(ByRef db As DAO.Database, ByVal tableName As String, ByVal userId As Long) As Boolean
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT FechaHPSConcesionMinima FROM " & tableName & " WHERE ID=" & userId, dbOpenSnapshot)
    If Not rs.EOF Then CacheDateIsNull = IsNull(rs!FechaHPSConcesionMinima)
    rs.Close
End Function

Private Function CardinalityPayload(ByVal beforeList As Long, ByVal beforeIndicators As Long, ByVal afterList As Long, ByVal afterIndicators As Long) As String
    CardinalityPayload = "{""beforeList"":" & beforeList & ",""beforeIndicators"":" & beforeIndicators & ",""afterList"":" & afterList & ",""afterIndicators"":" & afterIndicators & "}"
End Function

Private Function JsonOkPayload(ByVal value As String, ByVal payloadJson As String, ByRef logs As Collection) As String
    JsonOkPayload = "{""ok"":true,""value"":""" & EscapeJson(value) & """,""payload"":" & payloadJson & ",""error"":null,""logs"":" & LogsJson(logs) & "}"
End Function

Private Function JsonFail(ByVal message As String, ByRef logs As Collection) As String
    JsonFail = "{""ok"":false,""value"":null,""payload"":null,""error"":""" & EscapeJson(message) & """,""logs"":" & LogsJson(logs) & "}"
End Function

Private Function LogsJson(ByRef logs As Collection) As String
    Dim i As Long
    Dim result As String
    result = "["
    If Not logs Is Nothing Then
        For i = 1 To logs.Count
            If i > 1 Then result = result & ","
            result = result & """" & EscapeJson(CStr(logs(i))) & """"
        Next i
    End If
    LogsJson = result & "]"
End Function

Private Function EscapeJson(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, """", Chr$(92) & Chr$(34))
    value = Replace(value, vbCrLf, "\n")
    value = Replace(value, vbCr, "\n")
    value = Replace(value, vbLf, "\n")
    EscapeJson = value
End Function

