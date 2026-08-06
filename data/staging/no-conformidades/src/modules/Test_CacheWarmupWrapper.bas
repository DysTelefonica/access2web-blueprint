Attribute VB_Name = "Test_CacheWarmupWrapper"
Option Compare Database
Option Explicit

' ============================================================
' Test_CacheWarmupWrapper — tests TDD para el wrapper granular
' Cache_Warmup_Opciones (SDD: cache-idempotent-warmup-2026-06-17).
' ============================================================
' Verifica los 8 escenarios del spec:
'   1. CSV vacio -> procesa las 6 tiers, errores=[], resultados=6 entries.
'   2. Inclusion CSV -> solo las tiers listadas.
'   3. Exclusion CSV (-X,-Y) -> todas menos X e Y.
'   4. Dry-run no escribe (row counts before/after identicos).
'   5. EnsureSchema=True crea tablas audit faltantes.
'   6. EnsureSchema=False NO crea tablas; warmup falla gracefully.
'   7. Idempotente (2 llamadas -> mismo resultado, row counts estables).
'   8. JSON parseable por JsonConverter (round-trip).
'
' Fixture strategy: sandbox local via TestHelper.BeginTestSession +
' m_TestingMode=True. Setup con DELETE + (Tests 5/6) DROP, todos con
' On Error Resume Next para tolerar tablas inexistentes.
'
' IMPORTANTE: JsonConverter parsea JSON arrays como Collection (1-based),
' no como Variant() 0-based. Todos los accesos usan (1)..(Count).
' ============================================================

' ============================================================
' TEST 1 — CSV vacio puebla las 6 tiers
' ============================================================
Public Function Test_Cache_Warmup_Opciones_VacioPueblaTodo() As String
    On Error GoTo EH

    Dim logs As Collection
    Dim sessionErr As String
    Dim sessionStarted As Boolean
    Dim jsonReturn As String
    Dim parsed As Object
    Dim assertError As String
    Dim db As DAO.Database
    Dim dbErr As String

    Set logs = TestHelper.NewLogs
    sessionStarted = False

    If Not TestHelper.BeginTestSession(logs, sessionErr) Then
        Test_Cache_Warmup_Opciones_VacioPueblaTodo = TestHelper.BuildJsonFail("TESTS BLOCKED: " & sessionErr, logs)
        Exit Function
    End If
    sessionStarted = True

    Set db = getdb(dbErr)
    If db Is Nothing Then
        Test_Cache_Warmup_Opciones_VacioPueblaTodo = TestHelper.BuildJsonFail("getdb: " & dbErr, logs)
        GoTo Cleanup
    End If

    ' Setup: limpiar cache tables (idempotencia entre tests).
    ' On Error Resume Next: tablas pueden no existir en sandbox fresh.
    On Error Resume Next
    db.Execute "DELETE FROM TbCacheListadoNC"
    db.Execute "DELETE FROM TbCacheNCProyecto"
    db.Execute "DELETE FROM TbCacheListadoNCAuditoria"
    db.Execute "DELETE FROM TbEstadoCatalogo"
    db.Execute "DELETE FROM TbCacheIndicadoresProyectoDetalle"
    db.Execute "DELETE FROM TbCacheIndicadoresProyectoHeader"
    Err.Clear
    On Error GoTo EH

    ' Act: wrapper con CSV vacio.
    On Error Resume Next
    jsonReturn = Cache_Warmup_Opciones("", True, False, sessionErr)
    If Err.Number <> 0 Then
        Test_Cache_Warmup_Opciones_VacioPueblaTodo = TestHelper.BuildJsonFail("wrapper raised error: " & Err.Description, logs)
        On Error GoTo EH
        GoTo Cleanup
    End If
    On Error GoTo EH

    On Error Resume Next
    Set parsed = JsonConverter.ParseJson(jsonReturn)
    On Error GoTo EH
    If parsed Is Nothing Then
        Test_Cache_Warmup_Opciones_VacioPueblaTodo = TestHelper.BuildJsonFail("JSON no parseable: " & Left$(jsonReturn, 200), logs)
        GoTo Cleanup
    End If

    ' Assert 1: vias_procesadas tiene 6 tiers.
    If Not TestHelper.AssertTrue(CollectionCount(parsed("vias_procesadas")) = 6, _
                                  "Assert1: vias_procesadas debe tener 6 tiers; actual=" & CollectionCount(parsed("vias_procesadas")), _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_VacioPueblaTodo = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 2: vias_omitidas vacia.
    If Not TestHelper.AssertTrue(CollectionCount(parsed("vias_omitidas")) = 0, _
                                  "Assert2: vias_omitidas debe estar vacia; actual=" & CollectionCount(parsed("vias_omitidas")), _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_VacioPueblaTodo = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 3: errores vacio.
    If Not TestHelper.AssertTrue(CollectionCount(parsed("errores")) = 0, _
                                  "Assert3: errores debe estar vacio; actual=" & CollectionCount(parsed("errores")), _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_VacioPueblaTodo = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 4: 6 entradas en resultados.
    If Not TestHelper.AssertTrue(parsed("resultados").Count = 6, _
                                  "Assert4: resultados debe tener 6 entradas; actual=" & parsed("resultados").Count, _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_VacioPueblaTodo = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 5: dry_run=false y ensure_schema_ejecutado=true.
    If Not TestHelper.AssertTrue(CBool(parsed("dry_run")) = False, _
                                  "Assert5a: dry_run debe ser false; actual=" & parsed("dry_run"), _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_VacioPueblaTodo = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(CBool(parsed("ensure_schema_ejecutado")) = True, _
                                  "Assert5b: ensure_schema_ejecutado debe ser true; actual=" & parsed("ensure_schema_ejecutado"), _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_VacioPueblaTodo = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    Test_Cache_Warmup_Opciones_VacioPueblaTodo = TestHelper.BuildJsonOk(logs, "vacio_todo_ok")
    GoTo Cleanup

EH:
    TestHelper.AddLog logs, "Error: " & Err.Description
    Test_Cache_Warmup_Opciones_VacioPueblaTodo = TestHelper.BuildJsonFail("EH: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If sessionStarted Then Call TestHelper.EndTestSession(logs)
    On Error GoTo 0
End Function

' ============================================================
' TEST 2 — Inclusion CSV: solo LISTADO procesado
' ============================================================
Public Function Test_Cache_Warmup_Opciones_FiltrarPorInclusion() As String
    On Error GoTo EH

    Dim logs As Collection
    Dim sessionErr As String
    Dim sessionStarted As Boolean
    Dim jsonReturn As String
    Dim parsed As Object
    Dim assertError As String
    Dim db As DAO.Database
    Dim dbErr As String

    Set logs = TestHelper.NewLogs
    sessionStarted = False

    If Not TestHelper.BeginTestSession(logs, sessionErr) Then
        Test_Cache_Warmup_Opciones_FiltrarPorInclusion = TestHelper.BuildJsonFail("TESTS BLOCKED: " & sessionErr, logs)
        Exit Function
    End If
    sessionStarted = True

    Set db = getdb(dbErr)
    If db Is Nothing Then
        Test_Cache_Warmup_Opciones_FiltrarPorInclusion = TestHelper.BuildJsonFail("getdb: " & dbErr, logs)
        GoTo Cleanup
    End If

    On Error Resume Next
    db.Execute "DELETE FROM TbCacheListadoNC"
    Err.Clear
    On Error GoTo EH

    ' Act: CSV = "LISTADO".
    On Error Resume Next
    jsonReturn = Cache_Warmup_Opciones("LISTADO", True, False, sessionErr)
    If Err.Number <> 0 Then
        Test_Cache_Warmup_Opciones_FiltrarPorInclusion = TestHelper.BuildJsonFail("wrapper raised error: " & Err.Description, logs)
        On Error GoTo EH
        GoTo Cleanup
    End If
    On Error GoTo EH

    On Error Resume Next
    Set parsed = JsonConverter.ParseJson(jsonReturn)
    On Error GoTo EH
    If parsed Is Nothing Then
        Test_Cache_Warmup_Opciones_FiltrarPorInclusion = TestHelper.BuildJsonFail("JSON no parseable", logs)
        GoTo Cleanup
    End If

    ' Assert 1: vias_procesadas contiene solo LISTADO.
    If Not TestHelper.AssertTrue(CollectionCount(parsed("vias_procesadas")) = 1, _
                                  "Assert1: vias_procesadas debe tener 1 tier; actual=" & CollectionCount(parsed("vias_procesadas")), _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_FiltrarPorInclusion = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(UCase$(CStr(parsed("vias_procesadas")(1))) = "LISTADO", _
                                  "Assert1b: tier procesado debe ser LISTADO; actual=" & CStr(parsed("vias_procesadas")(1)), _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_FiltrarPorInclusion = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 2: vias_omitidas tiene las otras 5.
    If Not TestHelper.AssertTrue(CollectionCount(parsed("vias_omitidas")) = 5, _
                                  "Assert2: vias_omitidas debe tener 5 tiers; actual=" & CollectionCount(parsed("vias_omitidas")), _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_FiltrarPorInclusion = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 3: resultados solo tiene LISTADO.
    If Not TestHelper.AssertTrue(parsed("resultados").Count = 1, _
                                  "Assert3: resultados debe tener 1 entrada; actual=" & parsed("resultados").Count, _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_FiltrarPorInclusion = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(parsed("resultados").Exists("LISTADO"), _
                                  "Assert3b: resultados debe tener clave LISTADO", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_FiltrarPorInclusion = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    Test_Cache_Warmup_Opciones_FiltrarPorInclusion = TestHelper.BuildJsonOk(logs, "inclusion_ok")
    GoTo Cleanup

EH:
    TestHelper.AddLog logs, "Error: " & Err.Description
    Test_Cache_Warmup_Opciones_FiltrarPorInclusion = TestHelper.BuildJsonFail("EH: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If sessionStarted Then Call TestHelper.EndTestSession(logs)
    On Error GoTo 0
End Function

' ============================================================
' TEST 3 — Exclusion CSV: -X,-Y procesa el resto
' ============================================================
Public Function Test_Cache_Warmup_Opciones_FiltrarPorExclusion() As String
    On Error GoTo EH

    Dim logs As Collection
    Dim sessionErr As String
    Dim sessionStarted As Boolean
    Dim jsonReturn As String
    Dim parsed As Object
    Dim assertError As String
    Dim db As DAO.Database
    Dim dbErr As String
    Dim foundIndAud As Boolean
    Dim foundEstCat As Boolean
    Dim i As Long
    Dim omit As Variant
    Dim omitCol As Collection

    Set logs = TestHelper.NewLogs
    sessionStarted = False

    If Not TestHelper.BeginTestSession(logs, sessionErr) Then
        Test_Cache_Warmup_Opciones_FiltrarPorExclusion = TestHelper.BuildJsonFail("TESTS BLOCKED: " & sessionErr, logs)
        Exit Function
    End If
    sessionStarted = True

    Set db = getdb(dbErr)
    If db Is Nothing Then
        Test_Cache_Warmup_Opciones_FiltrarPorExclusion = TestHelper.BuildJsonFail("getdb: " & dbErr, logs)
        GoTo Cleanup
    End If

    On Error Resume Next
    db.Execute "DELETE FROM TbCacheListadoNC"
    db.Execute "DELETE FROM TbCacheNCProyecto"
    db.Execute "DELETE FROM TbCacheListadoNCAuditoria"
    db.Execute "DELETE FROM TbEstadoCatalogo"
    db.Execute "DELETE FROM TbCacheIndicadoresProyectoDetalle"
    db.Execute "DELETE FROM TbCacheIndicadoresProyectoHeader"
    Err.Clear
    On Error GoTo EH

    ' Act: CSV = "-INDICADORES_AUDITORIA,-ESTADO_CATALOGO".
    On Error Resume Next
    jsonReturn = Cache_Warmup_Opciones("-INDICADORES_AUDITORIA,-ESTADO_CATALOGO", True, False, sessionErr)
    If Err.Number <> 0 Then
        Test_Cache_Warmup_Opciones_FiltrarPorExclusion = TestHelper.BuildJsonFail("wrapper raised error: " & Err.Description, logs)
        On Error GoTo EH
        GoTo Cleanup
    End If
    On Error GoTo EH

    On Error Resume Next
    Set parsed = JsonConverter.ParseJson(jsonReturn)
    On Error GoTo EH
    If parsed Is Nothing Then
        Test_Cache_Warmup_Opciones_FiltrarPorExclusion = TestHelper.BuildJsonFail("JSON no parseable", logs)
        GoTo Cleanup
    End If

    ' Assert 1: vias_procesadas tiene 4 tiers (6 - 2).
    If Not TestHelper.AssertTrue(CollectionCount(parsed("vias_procesadas")) = 4, _
                                  "Assert1: vias_procesadas debe tener 4 tiers; actual=" & CollectionCount(parsed("vias_procesadas")), _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_FiltrarPorExclusion = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 2: vias_omitidas contiene INDICADORES_AUDITORIA y ESTADO_CATALOGO.
    Set omitCol = parsed("vias_omitidas")
    foundIndAud = False
    foundEstCat = False
    For i = 1 To omitCol.Count
        omit = CStr(omitCol(i))
        If StrComp(omit, "INDICADORES_AUDITORIA", vbTextCompare) = 0 Then foundIndAud = True
        If StrComp(omit, "ESTADO_CATALOGO", vbTextCompare) = 0 Then foundEstCat = True
    Next i
    If Not TestHelper.AssertTrue(foundIndAud And foundEstCat, _
                                  "Assert2: vias_omitidas debe contener INDICADORES_AUDITORIA y ESTADO_CATALOGO; indAud=" & foundIndAud & " estCat=" & foundEstCat, _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_FiltrarPorExclusion = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 3: vias_omitidas tiene exactamente 2 entries.
    If Not TestHelper.AssertTrue(omitCol.Count = 2, _
                                  "Assert3: vias_omitidas debe tener 2 entries; actual=" & omitCol.Count, _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_FiltrarPorExclusion = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 4: resultados tiene 4 entradas (no incluye INDICADORES_AUDITORIA ni ESTADO_CATALOGO).
    If Not TestHelper.AssertTrue(parsed("resultados").Count = 4, _
                                  "Assert4: resultados debe tener 4 entradas; actual=" & parsed("resultados").Count, _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_FiltrarPorExclusion = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(Not parsed("resultados").Exists("INDICADORES_AUDITORIA"), _
                                  "Assert4b: resultados NO debe incluir INDICADORES_AUDITORIA", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_FiltrarPorExclusion = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(Not parsed("resultados").Exists("ESTADO_CATALOGO"), _
                                  "Assert4c: resultados NO debe incluir ESTADO_CATALOGO", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_FiltrarPorExclusion = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    Test_Cache_Warmup_Opciones_FiltrarPorExclusion = TestHelper.BuildJsonOk(logs, "exclusion_ok")
    GoTo Cleanup

EH:
    TestHelper.AddLog logs, "Error: " & Err.Description
    Test_Cache_Warmup_Opciones_FiltrarPorExclusion = TestHelper.BuildJsonFail("EH: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If sessionStarted Then Call TestHelper.EndTestSession(logs)
    On Error GoTo 0
End Function

' ============================================================
' TEST 4 — Dry-run no escribe (row counts before/after identicos)
' ============================================================
Public Function Test_Cache_Warmup_Opciones_DryRunNoEscribe() As String
    On Error GoTo EH

    Dim logs As Collection
    Dim sessionErr As String
    Dim sessionStarted As Boolean
    Dim jsonReturn As String
    Dim parsed As Object
    Dim assertError As String
    Dim db As DAO.Database
    Dim dbErr As String
    Dim rowsListedBefore As Long
    Dim rowsListedAfter As Long
    Dim rowsDetalleBefore As Long
    Dim rowsDetalleAfter As Long
    Dim rowsIndicadoresProyBefore As Long
    Dim rowsIndicadoresProyAfter As Long
    Dim rowsEstadoBefore As Long
    Dim rowsEstadoAfter As Long
    Dim k As Variant
    Dim anyWould As Boolean

    Set logs = TestHelper.NewLogs
    sessionStarted = False

    If Not TestHelper.BeginTestSession(logs, sessionErr) Then
        Test_Cache_Warmup_Opciones_DryRunNoEscribe = TestHelper.BuildJsonFail("TESTS BLOCKED: " & sessionErr, logs)
        Exit Function
    End If
    sessionStarted = True

    Set db = getdb(dbErr)
    If db Is Nothing Then
        Test_Cache_Warmup_Opciones_DryRunNoEscribe = TestHelper.BuildJsonFail("getdb: " & dbErr, logs)
        GoTo Cleanup
    End If

    ' Setup: asegurar schema y limpiar cache tables.
    On Error Resume Next
    CacheNCProyecto.EnsureCacheSchemaReadiness sessionErr
    NCAuditoriaListadoCache.EnsureNCAuditoriaListadoCacheSchema sessionErr
    EstadoCatalogoBootstrap.EnsureEstadoCatalogoSchema db, sessionErr
    sessionErr = ""
    db.Execute "DELETE FROM TbCacheListadoNC"
    db.Execute "DELETE FROM TbCacheNCProyecto"
    db.Execute "DELETE FROM TbCacheListadoNCAuditoria"
    db.Execute "DELETE FROM TbEstadoCatalogo"
    db.Execute "DELETE FROM TbCacheIndicadoresProyectoDetalle"
    db.Execute "DELETE FROM TbCacheIndicadoresProyectoHeader"
    Err.Clear
    On Error GoTo EH

    ' Snapshot row counts BEFORE.
    rowsListedBefore = SafeCount(db, "SELECT COUNT(*) AS N FROM TbCacheListadoNC")
    rowsDetalleBefore = SafeCount(db, "SELECT COUNT(*) AS N FROM TbCacheNCProyecto")
    rowsIndicadoresProyBefore = SafeCount(db, "SELECT COUNT(*) AS N FROM TbCacheIndicadoresProyectoDetalle WHERE IDCacheIndicadorProyecto=1")
    rowsEstadoBefore = SafeCount(db, "SELECT COUNT(*) AS N FROM TbEstadoCatalogo")

    ' Act: dry-run.
    On Error Resume Next
    jsonReturn = Cache_Warmup_Opciones("", True, True, sessionErr)
    If Err.Number <> 0 Then
        Test_Cache_Warmup_Opciones_DryRunNoEscribe = TestHelper.BuildJsonFail("dry-run raised error: " & Err.Description, logs)
        On Error GoTo EH
        GoTo Cleanup
    End If
    On Error GoTo EH

    ' Snapshot row counts AFTER.
    rowsListedAfter = SafeCount(db, "SELECT COUNT(*) AS N FROM TbCacheListadoNC")
    rowsDetalleAfter = SafeCount(db, "SELECT COUNT(*) AS N FROM TbCacheNCProyecto")
    rowsIndicadoresProyAfter = SafeCount(db, "SELECT COUNT(*) AS N FROM TbCacheIndicadoresProyectoDetalle WHERE IDCacheIndicadorProyecto=1")
    rowsEstadoAfter = SafeCount(db, "SELECT COUNT(*) AS N FROM TbEstadoCatalogo")

    ' Assert 1: dry_run=true en JSON.
    On Error Resume Next
    Set parsed = JsonConverter.ParseJson(jsonReturn)
    On Error GoTo EH
    If parsed Is Nothing Then
        Test_Cache_Warmup_Opciones_DryRunNoEscribe = TestHelper.BuildJsonFail("JSON no parseable", logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(CBool(parsed("dry_run")) = True, _
                                  "Assert1: dry_run debe ser true; actual=" & parsed("dry_run"), _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_DryRunNoEscribe = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 2: row counts no cambian.
    If Not TestHelper.AssertTrue(rowsListedBefore = rowsListedAfter, _
                                  "Assert2a: TbCacheListadoNC row count estable (before=" & rowsListedBefore & " after=" & rowsListedAfter & ")", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_DryRunNoEscribe = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(rowsDetalleBefore = rowsDetalleAfter, _
                                  "Assert2b: TbCacheNCProyecto row count estable (before=" & rowsDetalleBefore & " after=" & rowsDetalleAfter & ")", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_DryRunNoEscribe = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(rowsIndicadoresProyBefore = rowsIndicadoresProyAfter, _
                                  "Assert2c: TbCacheIndicadoresProyectoDetalle Proyecto count estable (before=" & rowsIndicadoresProyBefore & " after=" & rowsIndicadoresProyAfter & ")", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_DryRunNoEscribe = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(rowsEstadoBefore = rowsEstadoAfter, _
                                  "Assert2d: TbEstadoCatalogo row count estable (before=" & rowsEstadoBefore & " after=" & rowsEstadoAfter & ")", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_DryRunNoEscribe = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 3: filas_que_insertaria presente en cada tier procesada (o al menos una).
    anyWould = False
    For Each k In parsed("resultados").Keys
        If parsed("resultados")(k).Exists("filas_que_insertaria") Then
            anyWould = True
            Exit For
        End If
    Next k
    If Not TestHelper.AssertTrue(anyWould, _
                                  "Assert3: al menos una tier debe reportar filas_que_insertaria en dry-run", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_DryRunNoEscribe = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    Test_Cache_Warmup_Opciones_DryRunNoEscribe = TestHelper.BuildJsonOk(logs, "dryrun_no_escribe_ok")
    GoTo Cleanup

EH:
    TestHelper.AddLog logs, "Error: " & Err.Description
    Test_Cache_Warmup_Opciones_DryRunNoEscribe = TestHelper.BuildJsonFail("EH: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If sessionStarted Then Call TestHelper.EndTestSession(logs)
    On Error GoTo 0
End Function

' ============================================================
' TEST 5 — EnsureSchema=True crea tablas audit faltantes
' ============================================================
Public Function Test_Cache_Warmup_Opciones_EnsureSchemaTrueCreaTablasFaltantes() As String
    On Error GoTo EH

    Dim logs As Collection
    Dim sessionErr As String
    Dim sessionStarted As Boolean
    Dim jsonReturn As String
    Dim parsed As Object
    Dim assertError As String
    Dim db As DAO.Database
    Dim dbErr As String

    Set logs = TestHelper.NewLogs
    sessionStarted = False

    If Not TestHelper.BeginTestSession(logs, sessionErr) Then
        Test_Cache_Warmup_Opciones_EnsureSchemaTrueCreaTablasFaltantes = TestHelper.BuildJsonFail("TESTS BLOCKED: " & sessionErr, logs)
        Exit Function
    End If
    sessionStarted = True

    Set db = getdb(dbErr)
    If db Is Nothing Then
        Test_Cache_Warmup_Opciones_EnsureSchemaTrueCreaTablasFaltantes = TestHelper.BuildJsonFail("getdb: " & dbErr, logs)
        GoTo Cleanup
    End If

    ' Setup: dropear audit tables (con On Error Resume Next — pueden no existir).
    ' Access DAO no tiene DROP TABLE IF EXISTS; swallow error 3376 (table missing).
    On Error Resume Next
    db.Execute "DROP TABLE TbCacheIndicadoresAuditoriaDetalle", dbFailOnError
    db.Execute "DROP TABLE TbCacheIndicadoresAuditoriaHeader", dbFailOnError
    Err.Clear
    On Error GoTo EH

    ' Verificar: ambas tablas deben estar missing antes del wrapper.
    ' Si una de las dos aun existe (caso patologico FK constraints), reportar y salir.
    If TableExistsHelper(db, "TbCacheIndicadoresAuditoriaHeader") Then
        Test_Cache_Warmup_Opciones_EnsureSchemaTrueCreaTablasFaltantes = TestHelper.BuildJsonFail("Setup invalido: TbCacheIndicadoresAuditoriaHeader aun existe (FK constraint?)", logs)
        GoTo Cleanup
    End If
    If TableExistsHelper(db, "TbCacheIndicadoresAuditoriaDetalle") Then
        Test_Cache_Warmup_Opciones_EnsureSchemaTrueCreaTablasFaltantes = TestHelper.BuildJsonFail("Setup invalido: TbCacheIndicadoresAuditoriaDetalle aun existe (FK constraint?)", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Setup OK: ambas audit tables missing antes del wrapper"

    ' Act: wrapper con EnsureSchema=True, solo INDICADORES_AUDITORIA.
    On Error Resume Next
    jsonReturn = Cache_Warmup_Opciones("INDICADORES_AUDITORIA", True, False, sessionErr)
    If Err.Number <> 0 Then
        Test_Cache_Warmup_Opciones_EnsureSchemaTrueCreaTablasFaltantes = TestHelper.BuildJsonFail("wrapper raised error: " & Err.Description, logs)
        On Error GoTo EH
        GoTo Cleanup
    End If
    On Error GoTo EH

    ' Refresh TableDefs: el wrapper abrio su propia conexion (m_Db) via getdb,
    ' distinta de esta db. CREATE TABLE via m_Db puede no ser visible en
    ' el TableDefs cacheado de db sin un refresh explicito.
    On Error Resume Next
    db.TableDefs.Refresh
    Err.Clear
    On Error GoTo EH

    ' Assert 1: ambas tablas ahora existen.
    If Not TestHelper.AssertTrue(TableExistsHelper(db, "TbCacheIndicadoresAuditoriaHeader"), _
                                  "Assert1a: TbCacheIndicadoresAuditoriaHeader debe existir despues del wrapper", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_EnsureSchemaTrueCreaTablasFaltantes = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(TableExistsHelper(db, "TbCacheIndicadoresAuditoriaDetalle"), _
                                  "Assert1b: TbCacheIndicadoresAuditoriaDetalle debe existir despues del wrapper", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_EnsureSchemaTrueCreaTablasFaltantes = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 2: ensure_schema_ejecutado=true en JSON.
    On Error Resume Next
    Set parsed = JsonConverter.ParseJson(jsonReturn)
    On Error GoTo EH
    If parsed Is Nothing Then
        Test_Cache_Warmup_Opciones_EnsureSchemaTrueCreaTablasFaltantes = TestHelper.BuildJsonFail("JSON no parseable", logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(CBool(parsed("ensure_schema_ejecutado")) = True, _
                                  "Assert2: ensure_schema_ejecutado debe ser true", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_EnsureSchemaTrueCreaTablasFaltantes = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 3: INDICADORES_AUDITORIA aparece en vias_procesadas y resultados.
    ' Spec contract: "ok=true is acceptable as long as the tables are present"
    ' — i.e., la creacion del schema es el verdadero outcome bajo test.
    ' El warmup puede o no insertar filas dependiendo de la data del sandbox,
    ' asi que NO assertamos ok=true estrictamente.
    If Not TestHelper.AssertTrue(parsed("resultados").Exists("INDICADORES_AUDITORIA"), _
                                  "Assert3: resultados debe incluir INDICADORES_AUDITORIA", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_EnsureSchemaTrueCreaTablasFaltantes = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    Test_Cache_Warmup_Opciones_EnsureSchemaTrueCreaTablasFaltantes = TestHelper.BuildJsonOk(logs, "ensure_schema_true_crea_ok")
    GoTo Cleanup

EH:
    TestHelper.AddLog logs, "Error: " & Err.Description
    Test_Cache_Warmup_Opciones_EnsureSchemaTrueCreaTablasFaltantes = TestHelper.BuildJsonFail("EH: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If sessionStarted Then Call TestHelper.EndTestSession(logs)
    On Error GoTo 0
End Function

' ============================================================
' TEST 6 — EnsureSchema=False NO crea tablas; warmup falla gracefully
' ============================================================
Public Function Test_Cache_Warmup_Opciones_EnsureSchemaFalseAsumeSchemaListo() As String
    On Error GoTo EH

    Dim logs As Collection
    Dim sessionErr As String
    Dim sessionStarted As Boolean
    Dim jsonReturn As String
    Dim parsed As Object
    Dim assertError As String
    Dim db As DAO.Database
    Dim dbErr As String

    Set logs = TestHelper.NewLogs
    sessionStarted = False

    If Not TestHelper.BeginTestSession(logs, sessionErr) Then
        Test_Cache_Warmup_Opciones_EnsureSchemaFalseAsumeSchemaListo = TestHelper.BuildJsonFail("TESTS BLOCKED: " & sessionErr, logs)
        Exit Function
    End If
    sessionStarted = True

    Set db = getdb(dbErr)
    If db Is Nothing Then
        Test_Cache_Warmup_Opciones_EnsureSchemaFalseAsumeSchemaListo = TestHelper.BuildJsonFail("getdb: " & dbErr, logs)
        GoTo Cleanup
    End If

    ' Setup: dropear audit tables.
    On Error Resume Next
    db.Execute "DROP TABLE TbCacheIndicadoresAuditoriaDetalle", dbFailOnError
    db.Execute "DROP TABLE TbCacheIndicadoresAuditoriaHeader", dbFailOnError
    Err.Clear
    On Error GoTo EH

    If TableExistsHelper(db, "TbCacheIndicadoresAuditoriaHeader") Or _
       TableExistsHelper(db, "TbCacheIndicadoresAuditoriaDetalle") Then
        Test_Cache_Warmup_Opciones_EnsureSchemaFalseAsumeSchemaListo = TestHelper.BuildJsonFail("Setup invalido: audit tables aun existen (FK constraint?)", logs)
        GoTo Cleanup
    End If
    TestHelper.AddLog logs, "Setup OK: ambas audit tables missing (simulating missing schema)"

    ' Act: wrapper con EnsureSchema=False, solo INDICADORES_AUDITORIA.
    On Error Resume Next
    jsonReturn = Cache_Warmup_Opciones("INDICADORES_AUDITORIA", False, False, sessionErr)
    If Err.Number <> 0 Then
        Test_Cache_Warmup_Opciones_EnsureSchemaFalseAsumeSchemaListo = TestHelper.BuildJsonFail("wrapper raised error: " & Err.Description, logs)
        On Error GoTo EH
        GoTo Cleanup
    End If
    On Error GoTo EH

    ' Assert 1: audit tables siguen missing (wrapper NO las creo).
    If Not TestHelper.AssertTrue(Not TableExistsHelper(db, "TbCacheIndicadoresAuditoriaHeader"), _
                                  "Assert1a: TbCacheIndicadoresAuditoriaHeader debe seguir missing (EnsureSchema=False)", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_EnsureSchemaFalseAsumeSchemaListo = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(Not TableExistsHelper(db, "TbCacheIndicadoresAuditoriaDetalle"), _
                                  "Assert1b: TbCacheIndicadoresAuditoriaDetalle debe seguir missing (EnsureSchema=False)", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_EnsureSchemaFalseAsumeSchemaListo = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 2: ensure_schema_ejecutado=false en JSON.
    On Error Resume Next
    Set parsed = JsonConverter.ParseJson(jsonReturn)
    On Error GoTo EH
    If parsed Is Nothing Then
        Test_Cache_Warmup_Opciones_EnsureSchemaFalseAsumeSchemaListo = TestHelper.BuildJsonFail("JSON no parseable", logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(CBool(parsed("ensure_schema_ejecutado")) = False, _
                                  "Assert2: ensure_schema_ejecutado debe ser false", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_EnsureSchemaFalseAsumeSchemaListo = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 3: warmup fallo gracefully (ok=false para INDICADORES_AUDITORIA).
    If Not TestHelper.AssertTrue(parsed("resultados").Exists("INDICADORES_AUDITORIA"), _
                                  "Assert3: resultados debe incluir INDICADORES_AUDITORIA", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_EnsureSchemaFalseAsumeSchemaListo = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(CBool(parsed("resultados")("INDICADORES_AUDITORIA")("ok")) = False, _
                                  "Assert3b: INDICADORES_AUDITORIA.ok debe ser false (warmup fallo gracefully)", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_EnsureSchemaFalseAsumeSchemaListo = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 4: errores contiene una entrada sobre audit tables faltantes.
    If Not TestHelper.AssertTrue(CollectionCount(parsed("errores")) >= 1, _
                                  "Assert4: errores debe tener >=1 entrada; actual=" & CollectionCount(parsed("errores")), _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_EnsureSchemaFalseAsumeSchemaListo = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    Test_Cache_Warmup_Opciones_EnsureSchemaFalseAsumeSchemaListo = TestHelper.BuildJsonOk(logs, "ensure_schema_false_falla_ok")
    GoTo Cleanup

EH:
    TestHelper.AddLog logs, "Error: " & Err.Description
    Test_Cache_Warmup_Opciones_EnsureSchemaFalseAsumeSchemaListo = TestHelper.BuildJsonFail("EH: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If sessionStarted Then Call TestHelper.EndTestSession(logs)
    On Error GoTo 0
End Function

' ============================================================
' TEST 7 — Idempotencia: dos llamadas -> mismo resultado
' ============================================================
Public Function Test_Cache_Warmup_Opciones_Idempotente() As String
    On Error GoTo EH

    Dim logs As Collection
    Dim sessionErr As String
    Dim sessionStarted As Boolean
    Dim jsonReturn1 As String
    Dim jsonReturn2 As String
    Dim parsed1 As Object
    Dim parsed2 As Object
    Dim assertError As String
    Dim db As DAO.Database
    Dim dbErr As String
    Dim rowsListed1 As Long
    Dim rowsListed2 As Long
    Dim rowsDetalle1 As Long
    Dim rowsDetalle2 As Long
    Dim rowsIndicadoresProy1 As Long
    Dim rowsIndicadoresProy2 As Long
    Dim rowsEstado1 As Long
    Dim rowsEstado2 As Long

    Set logs = TestHelper.NewLogs
    sessionStarted = False

    If Not TestHelper.BeginTestSession(logs, sessionErr) Then
        Test_Cache_Warmup_Opciones_Idempotente = TestHelper.BuildJsonFail("TESTS BLOCKED: " & sessionErr, logs)
        Exit Function
    End If
    sessionStarted = True

    Set db = getdb(dbErr)
    If db Is Nothing Then
        Test_Cache_Warmup_Opciones_Idempotente = TestHelper.BuildJsonFail("getdb: " & dbErr, logs)
        GoTo Cleanup
    End If

    ' Setup: asegurar schema y limpiar.
    On Error Resume Next
    CacheNCProyecto.EnsureCacheSchemaReadiness sessionErr
    NCAuditoriaListadoCache.EnsureNCAuditoriaListadoCacheSchema sessionErr
    EstadoCatalogoBootstrap.EnsureEstadoCatalogoSchema db, sessionErr
    sessionErr = ""
    db.Execute "DELETE FROM TbCacheListadoNC"
    db.Execute "DELETE FROM TbCacheNCProyecto"
    db.Execute "DELETE FROM TbCacheListadoNCAuditoria"
    db.Execute "DELETE FROM TbEstadoCatalogo"
    db.Execute "DELETE FROM TbCacheIndicadoresProyectoDetalle"
    db.Execute "DELETE FROM TbCacheIndicadoresProyectoHeader"
    Err.Clear
    On Error GoTo EH

    ' Primera llamada.
    On Error Resume Next
    jsonReturn1 = Cache_Warmup_Opciones("", True, False, sessionErr)
    If Err.Number <> 0 Then
        Test_Cache_Warmup_Opciones_Idempotente = TestHelper.BuildJsonFail("primera llamada raised error: " & Err.Description, logs)
        On Error GoTo EH
        GoTo Cleanup
    End If
    On Error GoTo EH

    ' Snapshot 1.
    rowsListed1 = SafeCount(db, "SELECT COUNT(*) AS N FROM TbCacheListadoNC")
    rowsDetalle1 = SafeCount(db, "SELECT COUNT(*) AS N FROM TbCacheNCProyecto")
    rowsIndicadoresProy1 = SafeCount(db, "SELECT COUNT(*) AS N FROM TbCacheIndicadoresProyectoDetalle WHERE IDCacheIndicadorProyecto=1")
    rowsEstado1 = SafeCount(db, "SELECT COUNT(*) AS N FROM TbEstadoCatalogo")

    ' Segunda llamada.
    On Error Resume Next
    jsonReturn2 = Cache_Warmup_Opciones("", True, False, sessionErr)
    If Err.Number <> 0 Then
        Test_Cache_Warmup_Opciones_Idempotente = TestHelper.BuildJsonFail("segunda llamada raised error: " & Err.Description, logs)
        On Error GoTo EH
        GoTo Cleanup
    End If
    On Error GoTo EH

    ' Snapshot 2.
    rowsListed2 = SafeCount(db, "SELECT COUNT(*) AS N FROM TbCacheListadoNC")
    rowsDetalle2 = SafeCount(db, "SELECT COUNT(*) AS N FROM TbCacheNCProyecto")
    rowsIndicadoresProy2 = SafeCount(db, "SELECT COUNT(*) AS N FROM TbCacheIndicadoresProyectoDetalle WHERE IDCacheIndicadorProyecto=1")
    rowsEstado2 = SafeCount(db, "SELECT COUNT(*) AS N FROM TbEstadoCatalogo")

    ' Assert 1: ambas JSON parseables.
    On Error Resume Next
    Set parsed1 = JsonConverter.ParseJson(jsonReturn1)
    Set parsed2 = JsonConverter.ParseJson(jsonReturn2)
    On Error GoTo EH
    If parsed1 Is Nothing Or parsed2 Is Nothing Then
        Test_Cache_Warmup_Opciones_Idempotente = TestHelper.BuildJsonFail("JSON no parseable en alguna llamada", logs)
        GoTo Cleanup
    End If

    ' Assert 2: row counts identicos despues de las 2 llamadas.
    If Not TestHelper.AssertTrue(rowsListed1 = rowsListed2, _
                                  "Assert2a: TbCacheListadoNC estable (call1=" & rowsListed1 & " call2=" & rowsListed2 & ")", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_Idempotente = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(rowsDetalle1 = rowsDetalle2, _
                                  "Assert2b: TbCacheNCProyecto estable (call1=" & rowsDetalle1 & " call2=" & rowsDetalle2 & ")", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_Idempotente = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(rowsIndicadoresProy1 = rowsIndicadoresProy2, _
                                  "Assert2c: TbCacheIndicadoresProyectoDetalle Proyecto estable (call1=" & rowsIndicadoresProy1 & " call2=" & rowsIndicadoresProy2 & ")", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_Idempotente = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(rowsEstado1 = rowsEstado2, _
                                  "Assert2d: TbEstadoCatalogo estable (call1=" & rowsEstado1 & " call2=" & rowsEstado2 & ")", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_Idempotente = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 3: ambas llamadas reportan la misma cantidad de tiers procesadas.
    If Not TestHelper.AssertTrue(CollectionCount(parsed1("vias_procesadas")) = CollectionCount(parsed2("vias_procesadas")), _
                                  "Assert3: vias_procesadas count estable (call1=" & CollectionCount(parsed1("vias_procesadas")) & " call2=" & CollectionCount(parsed2("vias_procesadas")) & ")", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_Idempotente = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    Test_Cache_Warmup_Opciones_Idempotente = TestHelper.BuildJsonOk(logs, "idempotente_ok")
    GoTo Cleanup

EH:
    TestHelper.AddLog logs, "Error: " & Err.Description
    Test_Cache_Warmup_Opciones_Idempotente = TestHelper.BuildJsonFail("EH: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If sessionStarted Then Call TestHelper.EndTestSession(logs)
    On Error GoTo 0
End Function

' ============================================================
' TEST 8 — JSON parseable: round-trip ParseJson + ConvertToJson
' ============================================================
Public Function Test_Cache_Warmup_Opciones_JsonEsValido() As String
    On Error GoTo EH

    Dim logs As Collection
    Dim sessionErr As String
    Dim sessionStarted As Boolean
    Dim jsonReturn As String
    Dim parsed As Object
    Dim rebuilt As String
    Dim reparsed As Object
    Dim assertError As String
    Dim db As DAO.Database
    Dim dbErr As String

    Set logs = TestHelper.NewLogs
    sessionStarted = False

    If Not TestHelper.BeginTestSession(logs, sessionErr) Then
        Test_Cache_Warmup_Opciones_JsonEsValido = TestHelper.BuildJsonFail("TESTS BLOCKED: " & sessionErr, logs)
        Exit Function
    End If
    sessionStarted = True

    Set db = getdb(dbErr)
    If db Is Nothing Then
        Test_Cache_Warmup_Opciones_JsonEsValido = TestHelper.BuildJsonFail("getdb: " & dbErr, logs)
        GoTo Cleanup
    End If

    On Error Resume Next
    CacheNCProyecto.EnsureCacheSchemaReadiness sessionErr
    sessionErr = ""
    On Error GoTo EH

    ' Act: invocar wrapper.
    On Error Resume Next
    jsonReturn = Cache_Warmup_Opciones("LISTADO", True, False, sessionErr)
    If Err.Number <> 0 Then
        Test_Cache_Warmup_Opciones_JsonEsValido = TestHelper.BuildJsonFail("wrapper raised error: " & Err.Description, logs)
        On Error GoTo EH
        GoTo Cleanup
    End If
    On Error GoTo EH

    ' Assert 1: ParseJson no devuelve Nothing.
    On Error Resume Next
    Set parsed = JsonConverter.ParseJson(jsonReturn)
    On Error GoTo EH
    If parsed Is Nothing Then
        Test_Cache_Warmup_Opciones_JsonEsValido = TestHelper.BuildJsonFail("ParseJson devolvio Nothing; raw=" & Left$(jsonReturn, 200), logs)
        GoTo Cleanup
    End If

    ' Assert 2: structural keys presentes.
    If Not TestHelper.AssertTrue(parsed.Exists("dry_run"), _
                                  "Assert2a: clave 'dry_run' presente", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_JsonEsValido = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(parsed.Exists("vias_procesadas"), _
                                  "Assert2b: clave 'vias_procesadas' presente", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_JsonEsValido = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(parsed.Exists("resultados"), _
                                  "Assert2c: clave 'resultados' presente", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_JsonEsValido = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(parsed.Exists("errores"), _
                                  "Assert2d: clave 'errores' presente", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_JsonEsValido = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If
    If Not TestHelper.AssertTrue(parsed.Exists("duracion_total_seg"), _
                                  "Assert2e: clave 'duracion_total_seg' presente", _
                                  logs, assertError) Then
        Test_Cache_Warmup_Opciones_JsonEsValido = TestHelper.BuildJsonFail(assertError, logs)
        GoTo Cleanup
    End If

    ' Assert 3: round-trip ConvertToJson no crashea.
    On Error Resume Next
    rebuilt = JsonConverter.ConvertToJson(parsed)
    On Error GoTo EH
    If Len(rebuilt) = 0 Then
        Test_Cache_Warmup_Opciones_JsonEsValido = TestHelper.BuildJsonFail("ConvertToJson devolvio string vacio", logs)
        GoTo Cleanup
    End If

    ' Assert 4: round-trip JSON re-parseable.
    On Error Resume Next
    Set reparsed = JsonConverter.ParseJson(rebuilt)
    On Error GoTo EH
    If reparsed Is Nothing Then
        Test_Cache_Warmup_Opciones_JsonEsValido = TestHelper.BuildJsonFail("round-trip JSON no re-parseable", logs)
        GoTo Cleanup
    End If
    If Not reparsed.Exists("vias_procesadas") Then
        Test_Cache_Warmup_Opciones_JsonEsValido = TestHelper.BuildJsonFail("round-trip JSON perdio clave 'vias_procesadas'", logs)
        GoTo Cleanup
    End If

    Test_Cache_Warmup_Opciones_JsonEsValido = TestHelper.BuildJsonOk(logs, "json_valido_ok")
    GoTo Cleanup

EH:
    TestHelper.AddLog logs, "Error: " & Err.Description
    Test_Cache_Warmup_Opciones_JsonEsValido = TestHelper.BuildJsonFail("EH: " & Err.Description, logs)

Cleanup:
    On Error Resume Next
    If sessionStarted Then Call TestHelper.EndTestSession(logs)
    On Error GoTo 0
End Function

' ============================================================
' HELPERS LOCALES — sandbox-safe wrappers para tests
' ============================================================

' SafeCount: COUNT(*) sandbox-safe (mismo patron que CountOnDb en
' InicializadorCache.bas). Si la tabla no existe o el query falla,
' devuelve 0. Usa el handle db pasado (no DCount, que apunta al
' CurrentDb del frontend linkeado y miente en tests).
Private Function SafeCount(ByRef p_Db As DAO.Database, ByVal p_Sql As String) As Long
    Dim rs As DAO.Recordset
    On Error GoTo notfound
    Set rs = p_Db.OpenRecordset(p_Sql, dbOpenSnapshot)
    If rs Is Nothing Then
        SafeCount = 0
        Exit Function
    End If
    If Not rs.EOF Then
        SafeCount = CLng(Nz(rs.Fields("N").Value, 0))
    Else
        SafeCount = 0
    End If
    rs.Close
    Set rs = Nothing
    On Error GoTo 0
    Exit Function
notfound:
    SafeCount = 0
    On Error GoTo 0
End Function

' TableExistsHelper: verifica si la tabla existe en el handle db.
' On Error Resume Next alrededor del TableDefs lookup (Access no
' tiene IF EXISTS).
Private Function TableExistsHelper(ByRef p_Db As DAO.Database, ByVal p_Tabla As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error GoTo notfound
    Set tdf = p_Db.TableDefs(p_Tabla)
    TableExistsHelper = True
    On Error GoTo 0
    Exit Function
notfound:
    TableExistsHelper = False
    On Error GoTo 0
End Function

' CollectionCount: helper defensivo para contar items de una estructura
' parseada desde JSON (Collection 1-based, Variant() array 0-based, o Empty/Null).
' Usa LenB + IIf + TypeName para tolerar los 3 shapes sin lanzar "Object required".
Private Function CollectionCount(ByVal p_Item As Variant) As Long
    On Error GoTo notfound
    If IsEmpty(p_Item) Or IsNull(p_Item) Then CollectionCount = 0: Exit Function
    If TypeName(p_Item) = "Collection" Then
        CollectionCount = p_Item.Count
    ElseIf TypeName(p_Item) = "Variant()" Or InStr(TypeName(p_Item), "()") > 0 Then
        CollectionCount = UBound(p_Item) - LBound(p_Item) + 1
    Else
        CollectionCount = 0
    End If
    On Error GoTo 0
    Exit Function
notfound:
    CollectionCount = 0
    On Error GoTo 0
End Function
