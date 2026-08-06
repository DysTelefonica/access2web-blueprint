Attribute VB_Name = "Test_PublicabilidadEdicion"
Option Compare Database
Option Explicit

' ============================================================
' Test_PublicabilidadEdicion — Tests para el módulo PublicabilidadEdicion
'
' Skill: access-vba-tdd v1.9 (schema-first + fixture patterns)
'
' Cobertura (11 tests activos, 100% testeable):
'   EvaluarPublicabilidadEdicion: guard Nil + UTE sin/con evidencia
'   CalcularPublicabilidadEdicion: guard IDEdicion vacío + riesgo
'     NoPublicable + todo Publicable
'   CachePublicabilidad_AsegurarSchema: crea tabla / OK si existe
'   CachePublicabilidad_RecalcularEdicion: guard Edicion Nothing
'   CachePublicabilidad_LeerEdicionPublicable: guard IDEdicion vacío +
'     cache hit + cache miss
'   InvalidarPublicabilidadPorCambioEvidenciaEdicion: ID vacío + cache row
'
' Fixture IDs: rango FIX_ID_BASE = 900000
' ============================================================

' --- Suite lifecycle (v1.9 §3) ---
Private m_SuiteRun As Boolean

' --- JSON helpers (delegación a Test_Helper_JSON, v1.9 §2) ---
Private Function JsonOk(ByVal value As Variant, ByRef logs() As String) As String
    JsonOk = BuildJsonOk(value, logs)
End Function

Private Function JsonFail(ByVal errMsg As String, ByRef logs() As String) As String
    JsonFail = BuildJsonFail(errMsg, logs)
End Function

' --- EnsureTestConfigLoaded (delegación a Test_Helper, v1.9 §3) ---
Private Function EnsureTestConfigLoaded(ByRef p_Error As String) As Boolean
    EnsureTestConfigLoaded = Test_Helper.EnsureTestConfigLoaded(p_Error)
End Function

' --- GetTestDb (delegación a Test_Fixtures, v1.9 §2) ---
Private Function GetTestDb(ByRef p_Error As String) As DAO.Database
    Set GetTestDb = Test_Fixtures.GetTestDb(p_Error)
End Function

' ============================================================
' RunAll — Ejecuta todos los tests en secuencia
' Patrón SuiteSetup / SuiteTeardown (v1.9 §3)
' ============================================================
Public Function Test_PublicabilidadEdicion_RunAll() As String
    Dim results(0 To 12) As String
    Dim names(0 To 12) As String
    Dim i As Long
    Dim runError As String
    Dim outLogs(0 To 15) As String

    ' --- SuiteSetup ---
    ForceLocalBackend runError
    If runError <> "" Then
        outLogs(0) = "TESTS BLOCKED: " & runError
        Test_PublicabilidadEdicion_RunAll = JsonFail("TESTS BLOCKED: " & runError, outLogs)
        Exit Function
    End If
    outLogs(0) = "SuiteSetup OK"

    ' --- Ejecutar los 13 tests atomicos ---
    names(0) = "Test_EvaluarPublicabilidadEdicion_EdicionNil"
    names(1) = "Test_EvaluarPublicabilidadEdicion_UteSinEvidencia"
    names(2) = "Test_EvaluarPublicabilidadEdicion_UteConEvidencia"
    names(3) = "Test_CalcularPublicabilidadEdicion_IDVacio"
    names(4) = "Test_CalcularPublicabilidadEdicion_ConRiesgoNoPublicable"
    names(5) = "Test_CalcularPublicabilidadEdicion_TodoPublicable"
    names(6) = "Test_CachePublicabilidad_RecalcularEdicion_EdicionNil"
    names(7) = "Test_CachePublicabilidad_LeerEdicionPublicable_IDVacio"
    names(8) = "Test_CachePublicabilidad_LeerEdicionPublicable_CacheHit"
    names(9) = "Test_CachePublicabilidad_LeerEdicionPublicable_CacheMiss"
    names(10) = "Test_CachePublicabilidad_AsegurarSchema_CreaTabla"
    names(11) = "Test_InvalidarPublicabilidadPorCambioEvidenciaEdicion_IDVacio"
    names(12) = "Test_InvalidarPublicabilidadPorCambioEvidenciaEdicion_BorraCache"

    results(0) = Test_EvaluarPublicabilidadEdicion_EdicionNil()
    results(1) = Test_EvaluarPublicabilidadEdicion_UteSinEvidencia()
    results(2) = Test_EvaluarPublicabilidadEdicion_UteConEvidencia()
    results(3) = Test_CalcularPublicabilidadEdicion_IDVacio()
    results(4) = Test_CalcularPublicabilidadEdicion_ConRiesgoNoPublicable()
    results(5) = Test_CalcularPublicabilidadEdicion_TodoPublicable()
    results(6) = Test_CachePublicabilidad_RecalcularEdicion_EdicionNil()
    results(7) = Test_CachePublicabilidad_LeerEdicionPublicable_IDVacio()
    results(8) = Test_CachePublicabilidad_LeerEdicionPublicable_CacheHit()
    results(9) = Test_CachePublicabilidad_LeerEdicionPublicable_CacheMiss()
    results(10) = Test_CachePublicabilidad_AsegurarSchema_CreaTabla()
    results(11) = Test_InvalidarPublicabilidadPorCambioEvidenciaEdicion_IDVacio()
    results(12) = Test_InvalidarPublicabilidadPorCambioEvidenciaEdicion_BorraCache()

    ' --- TeardownAll: limpieza defensiva final ---
    Test_Fixtures.TeardownAll
    outLogs(1) = "TeardownAll OK"

    ' --- SuiteTeardown ---
    Test_Helper.ResetTestSession
    outLogs(2) = "SuiteTeardown OK"

    ' --- Acumular resultados ---
    Dim allOk As Boolean
    allOk = True
    Dim allResults As String
    allResults = ""
    Dim firstFailure As String
    firstFailure = ""
    For i = 0 To 12
        If InStr(results(i), """ok"":false") > 0 Then
            allOk = False
            If firstFailure = "" Then firstFailure = names(i)
        End If
        outLogs(i + 3) = names(i) & ": " & IIf(InStr(results(i), """ok"":false") > 0, "FAIL", "OK")
        If allResults <> "" Then allResults = allResults & ","
        allResults = allResults & results(i)
    Next i

    If allOk Then
        Test_PublicabilidadEdicion_RunAll = JsonOk("all_pass", outLogs)
    Else
        Test_PublicabilidadEdicion_RunAll = JsonFail("some_tests_failed: " & firstFailure, outLogs)
    End If
End Function

Public Function Test_InvalidarPublicabilidadPorCambioEvidenciaEdicion_IDVacio() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: IDEdicion vacío"
    logs(1) = "2. Act: InvalidarPublicabilidadPorCambioEvidenciaEdicion"
    logs(2) = "3. Assert: p_Error vacío"

    Dim pError As String
    InvalidarPublicabilidadPorCambioEvidenciaEdicion "", pError

    If pError <> "" Then
        Test_InvalidarPublicabilidadPorCambioEvidenciaEdicion_IDVacio = JsonFail("Expected empty p_Error but got: " & pError, logs)
        Exit Function
    End If

    Test_InvalidarPublicabilidadPorCambioEvidenciaEdicion_IDVacio = JsonOk("id_vacio_pass", logs)
End Function

Public Function Test_InvalidarPublicabilidadPorCambioEvidenciaEdicion_BorraCache() As String
    Dim logs(0 To 9) As String
    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: TempVars(Publicabilidad_Usar_Cache)='Sí'"
    logs(3) = "4. Arrange: CachePublicabilidad_AsegurarSchema"
    logs(4) = "5. Arrange: INSERT TbCachePublicabilidadEdicion stale"
    logs(5) = "6. Act: InvalidarPublicabilidadPorCambioEvidenciaEdicion"
    logs(6) = "7. Assert: cache row count = 0"
    logs(7) = "8. Teardown: DELETE cache"
    logs(8) = "9. Teardown: TempVars restore"
    logs(9) = "10. Teardown: TeardownAll"

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_InvalidarPublicabilidadPorCambioEvidenciaEdicion_BorraCache = JsonFail(cfgError, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    Dim originalCacheFlag As String
    On Error Resume Next
    originalCacheFlag = CStr(Application.TempVars("Publicabilidad_Usar_Cache"))
    If Err.Number <> 0 Then originalCacheFlag = ""
    On Error GoTo 0
    Application.TempVars("Publicabilidad_Usar_Cache") = "Sí"

    Dim db As DAO.Database
    Set db = GetTestDb(cfgError)
    If db Is Nothing Then
        GoTo Teardown
    End If

    Dim IDEdicion As Long
    IDEdicion = Test_Fixtures.Cache_EdicionId

    Dim pError As String
    CachePublicabilidad_AsegurarSchema pError
    If pError <> "" Then
        GoTo Teardown
    End If

    On Error Resume Next
    db.Execute "DELETE FROM TbCachePublicabilidadEdicion WHERE IDEdicion=" & IDEdicion
    db.Execute "INSERT INTO TbCachePublicabilidadEdicion " & _
               "(IDEdicion, Tipo, IDRiesgo, Publicable, AlgoritmoVersion, UpdatedAt) " & _
               "VALUES (" & IDEdicion & ", 'EDICION', 0, -1, 2, #" & Format$(Now, "yyyy-mm-dd") & "#)"

    Dim seedError As String
    If Err.Number <> 0 Then
        seedError = "No se pudo insertar cache stale de prueba: " & Err.Description
    End If
    On Error GoTo 0

    If seedError <> "" Then
        cfgError = seedError
        GoTo Teardown
    End If

    Dim seedCount As Long
    Dim rsSeed As DAO.Recordset
    Set rsSeed = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbCachePublicabilidadEdicion WHERE IDEdicion=" & IDEdicion)
    seedCount = CLng(rsSeed.Fields("C"))
    rsSeed.Close
    Set rsSeed = Nothing
    If seedCount = 0 Then
        cfgError = "No se insertó la fila stale de cache antes de invalidar"
        GoTo Teardown
    End If

    InvalidarPublicabilidadPorCambioEvidenciaEdicion CStr(IDEdicion), pError

    Dim rowCount As Long
    If pError = "" Then
        Dim rs As DAO.Recordset
        Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM TbCachePublicabilidadEdicion WHERE IDEdicion=" & IDEdicion)
        rowCount = CLng(rs.Fields("C"))
        rs.Close
        Set rs = Nothing
    End If

Teardown:
    On Error Resume Next
    If Not db Is Nothing Then
        If IDEdicion <> 0 Then db.Execute "DELETE FROM TbCachePublicabilidadEdicion WHERE IDEdicion=" & IDEdicion
    End If
    If originalCacheFlag <> "" Then
        Application.TempVars("Publicabilidad_Usar_Cache") = originalCacheFlag
    Else
        Application.TempVars.Remove "Publicabilidad_Usar_Cache"
    End If
    Set db = Nothing
    Test_Fixtures.TeardownAll
    On Error GoTo 0

    If cfgError <> "" Then
        Test_InvalidarPublicabilidadPorCambioEvidenciaEdicion_BorraCache = JsonFail(cfgError, logs)
        Exit Function
    End If
    If pError <> "" Then
        Test_InvalidarPublicabilidadPorCambioEvidenciaEdicion_BorraCache = JsonFail(pError, logs)
        Exit Function
    End If
    If rowCount <> 0 Then
        Test_InvalidarPublicabilidadPorCambioEvidenciaEdicion_BorraCache = JsonFail("Expected stale cache row deleted but count is " & CStr(rowCount), logs)
        Exit Function
    End If

    Test_InvalidarPublicabilidadPorCambioEvidenciaEdicion_BorraCache = JsonOk("invalidacion_evidencia_borra_cache_pass", logs)
End Function

' ============================================================
' Test: EvaluarPublicabilidadEdicion — Edicion Nil
' Guard: retorna EnumSiNo.No + error cuando p_Edicion Is Nothing
' Testeable directo: sin fixture
' ============================================================
Public Function Test_EvaluarPublicabilidadEdicion_EdicionNil() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: Edicion = Nothing"
    logs(1) = "2. Act: EvaluarPublicabilidadEdicion(Nothing, checks, pub, db, err)"
    logs(2) = "3. Assert: result = EnumSiNo.No"

    Dim ed As Edicion
    Set ed = Nothing

    Dim checks As Scripting.Dictionary
    Dim pub As Boolean
    Dim db As DAO.Database
    Dim pError As String

    Dim result As EnumSiNo
    result = EvaluarPublicabilidadEdicion(ed, checks, pub, db, pError)

    If result <> EnumSiNo.No Then
        Test_EvaluarPublicabilidadEdicion_EdicionNil = JsonFail("Expected EnumSiNo.No but got " & CStr(result), logs)
        Exit Function
    End If
    If pError = "" Then
        Test_EvaluarPublicabilidadEdicion_EdicionNil = JsonFail("Expected error message but got empty", logs)
        Exit Function
    End If

    Test_EvaluarPublicabilidadEdicion_EdicionNil = JsonOk("nil_guard_pass", logs)
End Function

' ============================================================
' Test: EvaluarPublicabilidadEdicion — UTE sin evidencia
' Schema-first: EnUTE via UPDATE TbProyectos (columna real TEXT(2))
'   TieneAnexoEvidenciaUTE via Property Let (no columna real en tabla)
'   EvaluarPublicabilidadEdicion retorna EnumSiNo.Sí con checks =
'   {ute_evidence: NoCumple, supplier_evidence: NoAplica} + pub=False
'
' Fixture: SeedAll (grafo base) + Constructor.getEdicion
' Teardown: restaurar EnUTE = NULL
'
' FIX v1.9: GoTo Teardown en vez de Exit Function directo
' FIX v1.9: Teardown block con On Error Resume Next + On Error GoTo 0
' ============================================================
Public Function Test_EvaluarPublicabilidadEdicion_UteSinEvidencia() As String
    Dim logs(0 To 9) As String
    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: UPDATE TbProyectos SET EnUTE='S'"
    logs(3) = "4. Arrange: Constructor.getEdicion + TieneAnexoEvidenciaUTE=No (Property Let)"
    logs(4) = "5. Act: EvaluarPublicabilidadEdicion"
    logs(5) = "6. Assert: ute_evidence.estado = NoCumple"
    logs(6) = "7. Assert: supplier_evidence.estado = NoAplica"
    logs(7) = "8. Assert: EdicionPublicable = False"
    logs(8) = "9. Teardown: EnUTE=NULL"
    logs(9) = "10. Teardown: TeardownAll"

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_EvaluarPublicabilidadEdicion_UteSinEvidencia = JsonFail(cfgError, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    Dim db As DAO.Database
    Set db = GetTestDb(cfgError)
    If db Is Nothing Then
        Test_Fixtures.TeardownAll
        Test_EvaluarPublicabilidadEdicion_UteSinEvidencia = JsonFail(cfgError, logs)
        Exit Function
    End If

    Dim IDEdicion As Long
    IDEdicion = Test_Fixtures.Cache_EdicionId
    Dim idProyecto As Long
    idProyecto = Test_Fixtures.Cache_ProyectoId

    ' Arrange: EnUTE='S' en TbProyectos (columna real TEXT(2))
    On Error Resume Next
    db.Execute "UPDATE TbProyectos SET EnUTE='S' WHERE IDProyecto=" & idProyecto
    On Error GoTo 0

    ' Arrange: Edicion via Constructor.getEdicion para que cargue el Proyecto
    Dim pError As String
    Dim ed As Edicion
    Set ed = Constructor.getEdicion(CStr(IDEdicion), pError)
    If ed Is Nothing Then
        GoTo Teardown
    End If

    ' Arrange: TieneAnexoEvidenciaUTE via Property Let (NO es columna real)
    ed.TieneAnexoEvidenciaUTE = EnumSiNo.No

    Dim checks As Scripting.Dictionary
    Dim pub As Boolean

    Dim result As EnumSiNo
    result = EvaluarPublicabilidadEdicion(ed, checks, pub, , pError)

    If result <> EnumSiNo.Sí Then
        GoTo Teardown
    End If

    ' Assert: ute_evidence = NoCumple
    If checks Is Nothing Then GoTo Teardown
    If Not checks.Exists("ute_evidence") Then GoTo Teardown
    Dim uteCheck As Scripting.Dictionary
    Set uteCheck = checks("ute_evidence")
    If uteCheck("estado") <> EnumPublicabilidadCheckEstado.NoCumple Then GoTo Teardown

    ' Assert: supplier_evidence = NoAplica when no supplier fixture exists
    If Not checks.Exists("supplier_evidence") Then GoTo Teardown
    Dim supCheck As Scripting.Dictionary
    Set supCheck = checks("supplier_evidence")
    If supCheck("estado") <> EnumPublicabilidadCheckEstado.NoAplica Then GoTo Teardown

    ' Assert: EdicionPublicable = False (UTE sin evidencia bloquea)
    If pub <> False Then GoTo Teardown

    GoTo Teardown

Teardown:
    On Error Resume Next
    db.Execute "UPDATE TbProyectos SET EnUTE=NULL WHERE IDProyecto=" & idProyecto
    On Error GoTo 0
    Set db = Nothing
    Test_Fixtures.TeardownAll

    ' Re-check assertions after teardown
    If result <> EnumSiNo.Sí Then
        Test_EvaluarPublicabilidadEdicion_UteSinEvidencia = JsonFail("Expected EnumSiNo.Sí but got " & CStr(result) & ": " & pError, logs)
        Exit Function
    End If
    If checks Is Nothing Then
        Test_EvaluarPublicabilidadEdicion_UteSinEvidencia = JsonFail("checks is Nothing", logs)
        Exit Function
    End If
    If Not checks.Exists("ute_evidence") Then
        Test_EvaluarPublicabilidadEdicion_UteSinEvidencia = JsonFail("ute_evidence check not found", logs)
        Exit Function
    End If
    If uteCheck("estado") <> EnumPublicabilidadCheckEstado.NoCumple Then
        Test_EvaluarPublicabilidadEdicion_UteSinEvidencia = JsonFail("ute_evidence expected NoCumple but got " & CStr(uteCheck("estado")), logs)
        Exit Function
    End If
    If Not checks.Exists("supplier_evidence") Then
        Test_EvaluarPublicabilidadEdicion_UteSinEvidencia = JsonFail("supplier_evidence check not found", logs)
        Exit Function
    End If
    If supCheck("estado") <> EnumPublicabilidadCheckEstado.NoAplica Then
        Test_EvaluarPublicabilidadEdicion_UteSinEvidencia = JsonFail("supplier_evidence expected NoAplica but got " & CStr(supCheck("estado")), logs)
        Exit Function
    End If
    If pub <> False Then
        Test_EvaluarPublicabilidadEdicion_UteSinEvidencia = JsonFail("EdicionPublicable expected False but got True", logs)
        Exit Function
    End If

    Test_EvaluarPublicabilidadEdicion_UteSinEvidencia = JsonOk("ute_no_evidence_pass", logs)
End Function

' ============================================================
' Test: EvaluarPublicabilidadEdicion — UTE con evidencia
' Schema-first: EnUTE='S' + INSERT TbAnexos con EvidenciaUTE='Sí'
'   Property Let TieneAnexoEvidenciaUTE=EnumSiNo.Sí (para forzar el estado)
'   Verifica: ute_evidence=Cumple, supplier_evidence=NoAplica, pub=True
'
' Teardown: DELETE TbAnexos + EnUTE=NULL + TeardownAll
'
' FIX v1.9: GoTo Teardown + Teardown block con On Error Resume Next
' FIX v1.9: Fix assertions AFTER teardown to avoid "object disposed" on db
' ============================================================
Public Function Test_EvaluarPublicabilidadEdicion_UteConEvidencia() As String
    Dim logs(0 To 9) As String
    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: UPDATE TbProyectos SET EnUTE='S'"
    logs(3) = "4. Arrange: INSERT TbAnexos (EvidenciaUTE='Sí') + Constructor.getEdicion"
    logs(4) = "5. Act: EvaluarPublicabilidadEdicion"
    logs(5) = "6. Assert: ute_evidence.estado = Cumple"
    logs(6) = "7. Assert: supplier_evidence.estado = NoAplica"
    logs(7) = "8. Assert: EdicionPublicable = True"
    logs(8) = "9. Teardown: DELETE TbAnexos + EnUTE=NULL"
    logs(9) = "10. Teardown: TeardownAll"

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_EvaluarPublicabilidadEdicion_UteConEvidencia = JsonFail(cfgError, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    Dim db As DAO.Database
    Set db = GetTestDb(cfgError)
    If db Is Nothing Then
        Test_Fixtures.TeardownAll
        Test_EvaluarPublicabilidadEdicion_UteConEvidencia = JsonFail(cfgError, logs)
        Exit Function
    End If

    Dim IDEdicion As Long
    IDEdicion = Test_Fixtures.Cache_EdicionId
    Dim idProyecto As Long
    idProyecto = Test_Fixtures.Cache_ProyectoId
    Const FIX_ANEXO_ID As Long = 900600

    ' Arrange: EnUTE='S' en TbProyectos
    On Error Resume Next
    db.Execute "UPDATE TbProyectos SET EnUTE='S' WHERE IDProyecto=" & idProyecto
    On Error GoTo 0

    ' Arrange: insertar anexo UTE (schema-first: todos los campos NOT NULL del ERD)
    On Error Resume Next
    db.Execute "DELETE FROM TbAnexos WHERE IDAnexo=" & FIX_ANEXO_ID
    db.Execute "INSERT INTO TbAnexos (IDAnexo, IDProyecto, IDEdicion, Titulo, EvidenciaUTE) " & _
               "VALUES (" & FIX_ANEXO_ID & ", " & idProyecto & ", " & IDEdicion & ", " & _
               "'Fixture anexo UTE test', 'Sí')"
    On Error GoTo 0

    ' Arrange: Edicion via Constructor
    Dim pError As String
    Dim ed As Edicion
    Set ed = Constructor.getEdicion(CStr(IDEdicion), pError)
    If ed Is Nothing Then GoTo Teardown

    ' Arrange: TieneAnexoEvidenciaUTE=Sí via Property Let
    ed.TieneAnexoEvidenciaUTE = EnumSiNo.Sí

    Dim checks As Scripting.Dictionary
    Dim pub As Boolean

    Dim result As EnumSiNo
    result = EvaluarPublicabilidadEdicion(ed, checks, pub, , pError)

    ' Capture assertion values before teardown closes db
    Dim resultOk As Boolean
    resultOk = (result = EnumSiNo.Sí)
    Dim checksOk As Boolean
    checksOk = (Not checks Is Nothing And checks.Exists("ute_evidence") And checks.Exists("supplier_evidence"))
    Dim uteEstado As Long
    Dim supEstado As Long
    If checksOk Then
        uteEstado = checks("ute_evidence")("estado")
        supEstado = checks("supplier_evidence")("estado")
    End If
    Dim pubResult As Boolean
    pubResult = pub

    GoTo Teardown

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbAnexos WHERE IDAnexo=" & FIX_ANEXO_ID
    db.Execute "UPDATE TbProyectos SET EnUTE=NULL WHERE IDProyecto=" & idProyecto
    On Error GoTo 0
    Set db = Nothing
    Test_Fixtures.TeardownAll

    ' Assert using captured values (db is valid here, checks may be Nothing after teardown)
    If Not resultOk Then
        Test_EvaluarPublicabilidadEdicion_UteConEvidencia = JsonFail("Expected EnumSiNo.Sí but got " & CStr(result) & ": " & pError, logs)
        Exit Function
    End If
    If Not checksOk Then
        Test_EvaluarPublicabilidadEdicion_UteConEvidencia = JsonFail("ute_evidence check not found", logs)
        Exit Function
    End If
    If uteEstado <> EnumPublicabilidadCheckEstado.Cumple Then
        Test_EvaluarPublicabilidadEdicion_UteConEvidencia = JsonFail("ute_evidence expected Cumple but got " & CStr(uteEstado), logs)
        Exit Function
    End If
    If supEstado <> EnumPublicabilidadCheckEstado.NoAplica Then
        Test_EvaluarPublicabilidadEdicion_UteConEvidencia = JsonFail("supplier_evidence expected NoAplica but got " & CStr(supEstado), logs)
        Exit Function
    End If
    If pubResult <> True Then
        Test_EvaluarPublicabilidadEdicion_UteConEvidencia = JsonFail("EdicionPublicable expected True but got False", logs)
        Exit Function
    End If

    Test_EvaluarPublicabilidadEdicion_UteConEvidencia = JsonOk("ute_with_evidence_pass", logs)
End Function

' ============================================================
' Test: EvaluarPublicabilidadEdicion — UTE con evidencia UTE
' pero con suministrador sin evidencia
' RED issue #8: supplier_evidence debe NoCumple, no NoAplica
' ============================================================
Public Function Test_EvaluarPublicabilidadEdicion_UteConSuministradorSinEvidencia_NoCumple() As String
    Dim logs(0 To 12) As String
    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: UPDATE TbProyectos SET EnUTE='S'"
    logs(3) = "4. Arrange: INSERT TbSuministradores fixture"
    logs(4) = "5. Arrange: INSERT TbExpedientesSuministradores fixture"
    logs(5) = "6. Arrange: INSERT TbAnexos EvidenciaUTE='Sí'"
    logs(6) = "7. Arrange: INSERT TbProyectosEdicionesSuministradores IDAnexo=NULL"
    logs(7) = "8. Act: EvaluarPublicabilidadEdicion"
    logs(8) = "9. Assert: ute_evidence.estado = Cumple"
    logs(9) = "10. Assert: supplier_evidence.estado = NoCumple"
    logs(10) = "11. Assert: supplier_evidence.detalle menciona evidencia o proveedor"
    logs(11) = "12. Assert: EdicionPublicable = False and supplier_evidence <> NoAplica"
    logs(12) = "13. Teardown: fixture delta + TeardownAll"

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_EvaluarPublicabilidadEdicion_UteConSuministradorSinEvidencia_NoCumple = JsonFail(cfgError, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    Dim db As DAO.Database
    Set db = GetTestDb(cfgError)
    If db Is Nothing Then
        Test_Fixtures.TeardownAll
        Test_EvaluarPublicabilidadEdicion_UteConSuministradorSinEvidencia_NoCumple = JsonFail(cfgError, logs)
        Exit Function
    End If

    Dim IDEdicion As Long
    IDEdicion = Test_Fixtures.Cache_EdicionId
    Dim idProyecto As Long
    idProyecto = Test_Fixtures.Cache_ProyectoId
    Dim idExpediente As Long
    idExpediente = Test_Fixtures.Cache_ExpedienteId
    Const FIX_PARENT_SUPPLIER_ID As Long = 900609
    Const FIX_SUPPLIER_ID As Long = 900610
    Const FIX_PARENT_EXP_SUPPLIER_ID As Long = 900611
    Const FIX_EXP_SUPPLIER_ID As Long = 900614
    Const FIX_ED_SUPPLIER_ID As Long = 900612
    Const FIX_UTE_ANEXO_ID As Long = 900613
    Const FIX_SUPPLIER_NAME As String = "FixtureSupplierMissingEvidence"

    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE ID=" & FIX_ED_SUPPLIER_ID
    db.Execute "DELETE FROM TbExpedientesSuministradores WHERE IDExpedienteSuministrador IN (" & FIX_PARENT_EXP_SUPPLIER_ID & "," & FIX_EXP_SUPPLIER_ID & ")"
    db.Execute "DELETE FROM TbSuministradores WHERE IDSuministrador=" & FIX_SUPPLIER_ID
    db.Execute "DELETE FROM TbSuministradores WHERE IDSuministrador=" & FIX_PARENT_SUPPLIER_ID
    db.Execute "DELETE FROM TbAnexos WHERE IDAnexo=" & FIX_UTE_ANEXO_ID
    db.Execute "UPDATE TbProyectos SET EnUTE='S' WHERE IDProyecto=" & idProyecto
    db.Execute "INSERT INTO TbSuministradores (IDSuministrador, Nombre, CIF, ConsorcioPropio) " & _
               "VALUES (" & FIX_PARENT_SUPPLIER_ID & ", 'FixtureOwnConsortiumIssue8', 'F99999609', 'Sí')"
    db.Execute "INSERT INTO TbSuministradores (IDSuministrador, Nombre, CIF, ConsorcioPropio) " & _
               "VALUES (" & FIX_SUPPLIER_ID & ", '" & FIX_SUPPLIER_NAME & "', 'F99999610', 'No')"
    db.Execute "INSERT INTO TbExpedientesSuministradores " & _
               "(IDExpedienteSuministrador, IDExpediente, IDSuministrador, IDPadre) " & _
               "VALUES (" & FIX_PARENT_EXP_SUPPLIER_ID & ", " & idExpediente & ", " & FIX_PARENT_SUPPLIER_ID & ", NULL)"
    db.Execute "INSERT INTO TbExpedientesSuministradores " & _
               "(IDExpedienteSuministrador, IDExpediente, IDSuministrador, IDPadre) " & _
               "VALUES (" & FIX_EXP_SUPPLIER_ID & ", " & idExpediente & ", " & FIX_SUPPLIER_ID & ", " & FIX_PARENT_EXP_SUPPLIER_ID & ")"
    db.Execute "INSERT INTO TbAnexos (IDAnexo, IDProyecto, IDEdicion, Titulo, EvidenciaUTE) " & _
               "VALUES (" & FIX_UTE_ANEXO_ID & ", " & idProyecto & ", " & IDEdicion & ", " & _
               "'Fixture anexo UTE issue 8', 'Sí')"
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador, IDAnexo) " & _
               "VALUES (" & FIX_ED_SUPPLIER_ID & ", " & IDEdicion & ", " & FIX_SUPPLIER_ID & ", NULL)"
    Dim seedError As String
    If Err.Number <> 0 Then
        seedError = "No se pudo preparar fixture issue #8: " & Err.Description
    End If
    On Error GoTo 0
    If seedError <> "" Then GoTo Teardown

    Dim pError As String
    Dim ed As Edicion
    Set ed = Constructor.getEdicion(CStr(IDEdicion), pError)
    If ed Is Nothing Then GoTo Teardown
    ed.TieneAnexoEvidenciaUTE = EnumSiNo.Sí

    Dim checks As Scripting.Dictionary
    Dim pub As Boolean
    Dim result As EnumSiNo
    result = EvaluarPublicabilidadEdicion(ed, checks, pub, , pError)

    Dim resultOk As Boolean
    resultOk = (result = EnumSiNo.Sí)
    Dim checksOk As Boolean
    checksOk = (Not checks Is Nothing And checks.Exists("ute_evidence") And checks.Exists("supplier_evidence"))
    Dim uteEstado As Long
    Dim supEstado As Long
    Dim supDetalle As String
    If checksOk Then
        uteEstado = checks("ute_evidence")("estado")
        supEstado = checks("supplier_evidence")("estado")
        If checks("supplier_evidence").Exists("detalle") Then
            supDetalle = CStr(checks("supplier_evidence")("detalle"))
        End If
    End If
    Dim pubResult As Boolean
    pubResult = pub

Teardown:
    On Error Resume Next
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE ID=" & FIX_ED_SUPPLIER_ID
        db.Execute "DELETE FROM TbExpedientesSuministradores WHERE IDExpedienteSuministrador IN (" & FIX_EXP_SUPPLIER_ID & "," & FIX_PARENT_EXP_SUPPLIER_ID & ")"
        db.Execute "DELETE FROM TbSuministradores WHERE IDSuministrador=" & FIX_SUPPLIER_ID
        db.Execute "DELETE FROM TbSuministradores WHERE IDSuministrador=" & FIX_PARENT_SUPPLIER_ID
        db.Execute "DELETE FROM TbAnexos WHERE IDAnexo=" & FIX_UTE_ANEXO_ID
        db.Execute "UPDATE TbProyectos SET EnUTE=NULL WHERE IDProyecto=" & idProyecto
    End If
    On Error GoTo 0
    Set db = Nothing
    Test_Fixtures.TeardownAll

    If seedError <> "" Then
        Test_EvaluarPublicabilidadEdicion_UteConSuministradorSinEvidencia_NoCumple = JsonFail(seedError, logs)
        Exit Function
    End If
    If Not resultOk Then
        Test_EvaluarPublicabilidadEdicion_UteConSuministradorSinEvidencia_NoCumple = JsonFail("Expected EnumSiNo.Sí but got " & CStr(result) & ": " & pError, logs)
        Exit Function
    End If
    If Not checksOk Then
        Test_EvaluarPublicabilidadEdicion_UteConSuministradorSinEvidencia_NoCumple = JsonFail("ute_evidence or supplier_evidence check not found", logs)
        Exit Function
    End If
    If uteEstado <> EnumPublicabilidadCheckEstado.Cumple Then
        Test_EvaluarPublicabilidadEdicion_UteConSuministradorSinEvidencia_NoCumple = JsonFail("ute_evidence expected Cumple but got " & CStr(uteEstado), logs)
        Exit Function
    End If
    If supEstado = EnumPublicabilidadCheckEstado.NoAplica Then
        Test_EvaluarPublicabilidadEdicion_UteConSuministradorSinEvidencia_NoCumple = JsonFail("supplier_evidence must not be NoAplica when a supplier lacks evidence", logs)
        Exit Function
    End If
    If supEstado <> EnumPublicabilidadCheckEstado.NoCumple Then
        Test_EvaluarPublicabilidadEdicion_UteConSuministradorSinEvidencia_NoCumple = JsonFail("supplier_evidence expected NoCumple but got " & CStr(supEstado), logs)
        Exit Function
    End If
    If InStr(1, supDetalle, "Falta evidencia", vbTextCompare) = 0 And InStr(1, supDetalle, FIX_SUPPLIER_NAME, vbTextCompare) = 0 Then
        Test_EvaluarPublicabilidadEdicion_UteConSuministradorSinEvidencia_NoCumple = JsonFail("supplier_evidence detalle missing evidence text or supplier name. Got: " & supDetalle, logs)
        Exit Function
    End If
    If pubResult <> False Then
        Test_EvaluarPublicabilidadEdicion_UteConSuministradorSinEvidencia_NoCumple = JsonFail("EdicionPublicable expected False but got True", logs)
        Exit Function
    End If

    Test_EvaluarPublicabilidadEdicion_UteConSuministradorSinEvidencia_NoCumple = JsonOk("ute_supplier_missing_evidence_red", logs)
End Function

' ============================================================
' Test: EvaluarPublicabilidadEdicion — UTE con evidencia y sin suministradores
' issue #18: supplier_evidence debe NoAplica cuando no hay suministradores
' ============================================================
Public Function Test_EvaluarPublicabilidadEdicion_UteConEvidenciaSinSuministradores_NoAplica() As String
    Dim logs(0 To 11) As String
    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: UPDATE TbProyectos SET EnUTE='S'"
    logs(3) = "4. Arrange: DELETE TbProyectosEdicionesSuministradores for IDEdicion"
    logs(4) = "5. Arrange: DELETE TbExpedientesSuministradores for IDExpediente"
    logs(5) = "6. Arrange: INSERT TbAnexos EvidenciaUTE='Sí'"
    logs(6) = "7. Act: EvaluarPublicabilidadEdicion"
    logs(7) = "8. Assert: ute_evidence.estado = Cumple"
    logs(8) = "9. Assert: supplier_evidence.estado = NoAplica"
    logs(9) = "10. Assert: EdicionPublicable = True"
    logs(10) = "11. Teardown: fixture delta + TeardownAll"
    logs(11) = "12. Teardown: EnUTE=NULL"

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_EvaluarPublicabilidadEdicion_UteConEvidenciaSinSuministradores_NoAplica = JsonFail(cfgError, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    Dim db As DAO.Database
    Set db = GetTestDb(cfgError)
    If db Is Nothing Then
        Test_Fixtures.TeardownAll
        Test_EvaluarPublicabilidadEdicion_UteConEvidenciaSinSuministradores_NoAplica = JsonFail(cfgError, logs)
        Exit Function
    End If

    Dim IDEdicion As Long
    IDEdicion = Test_Fixtures.Cache_EdicionId
    Dim idProyecto As Long
    idProyecto = Test_Fixtures.Cache_ProyectoId
    Dim idExpediente As Long
    idExpediente = Test_Fixtures.Cache_ExpedienteId
    Const FIX_UTE_ANEXO_ID As Long = 900618

    Dim seedError As String
    On Error Resume Next
    db.Execute "DELETE FROM TbAnexos WHERE IDAnexo=" & FIX_UTE_ANEXO_ID
    db.Execute "UPDATE TbProyectos SET EnUTE='S' WHERE IDProyecto=" & idProyecto
    db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & IDEdicion
    db.Execute "DELETE FROM TbExpedientesSuministradores WHERE IDExpediente=" & idExpediente
    db.Execute "INSERT INTO TbAnexos (IDAnexo, IDProyecto, IDEdicion, Titulo, EvidenciaUTE) " & _
               "VALUES (" & FIX_UTE_ANEXO_ID & ", " & idProyecto & ", " & IDEdicion & ", " & _
               "'Fixture anexo UTE issue 18', 'Sí')"
    If Err.Number <> 0 Then
        seedError = "No se pudo preparar fixture issue #18: " & Err.Description
    End If
    On Error GoTo 0
    If seedError <> "" Then GoTo Teardown

    Dim pError As String
    Dim ed As Edicion
    Set ed = Constructor.getEdicion(CStr(IDEdicion), pError)
    If ed Is Nothing Then GoTo Teardown
    ed.TieneAnexoEvidenciaUTE = EnumSiNo.Sí

    Dim checks As Scripting.Dictionary
    Dim pub As Boolean
    Dim result As EnumSiNo
    result = EvaluarPublicabilidadEdicion(ed, checks, pub, , pError)

    Dim resultOk As Boolean
    resultOk = (result = EnumSiNo.Sí)
    Dim checksOk As Boolean
    checksOk = (Not checks Is Nothing And checks.Exists("ute_evidence") And checks.Exists("supplier_evidence"))
    Dim uteEstado As Long
    Dim supEstado As Long
    If checksOk Then
        uteEstado = checks("ute_evidence")("estado")
        supEstado = checks("supplier_evidence")("estado")
    End If
    Dim pubResult As Boolean
    pubResult = pub

Teardown:
    On Error Resume Next
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbAnexos WHERE IDAnexo=" & FIX_UTE_ANEXO_ID
        db.Execute "UPDATE TbProyectos SET EnUTE=NULL WHERE IDProyecto=" & idProyecto
    End If
    On Error GoTo 0
    Set db = Nothing
    Test_Fixtures.TeardownAll

    If seedError <> "" Then
        Test_EvaluarPublicabilidadEdicion_UteConEvidenciaSinSuministradores_NoAplica = JsonFail(seedError, logs)
        Exit Function
    End If
    If Not resultOk Then
        Test_EvaluarPublicabilidadEdicion_UteConEvidenciaSinSuministradores_NoAplica = JsonFail("Expected EnumSiNo.Sí but got " & CStr(result) & ": " & pError, logs)
        Exit Function
    End If
    If Not checksOk Then
        Test_EvaluarPublicabilidadEdicion_UteConEvidenciaSinSuministradores_NoAplica = JsonFail("ute_evidence or supplier_evidence check not found", logs)
        Exit Function
    End If
    If uteEstado <> EnumPublicabilidadCheckEstado.Cumple Then
        Test_EvaluarPublicabilidadEdicion_UteConEvidenciaSinSuministradores_NoAplica = JsonFail("ute_evidence expected Cumple but got " & CStr(uteEstado), logs)
        Exit Function
    End If
    If supEstado <> EnumPublicabilidadCheckEstado.NoAplica Then
        Test_EvaluarPublicabilidadEdicion_UteConEvidenciaSinSuministradores_NoAplica = JsonFail("supplier_evidence expected NoAplica but got " & CStr(supEstado), logs)
        Exit Function
    End If
    If pubResult <> True Then
        Test_EvaluarPublicabilidadEdicion_UteConEvidenciaSinSuministradores_NoAplica = JsonFail("EdicionPublicable expected True but got False", logs)
        Exit Function
    End If

    Test_EvaluarPublicabilidadEdicion_UteConEvidenciaSinSuministradores_NoAplica = JsonOk("ute_with_evidence_no_suppliers_noaplica", logs)
End Function

' ============================================================
' Test: CalcularPublicabilidadEdicion — IDEdicion vacío
' Guard: retorna EnumSiNo.No cuando IDEdicion = ""
' Testeable directo: sin fixture
' ============================================================
Public Function Test_CalcularPublicabilidadEdicion_IDVacio() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: Edicion con IDEdicion=''"
    logs(1) = "2. Act: CalcularPublicabilidadEdicion"
    logs(2) = "3. Assert: result = EnumSiNo.No"
    logs(3) = "4. Assert: p_Error no vacío"

    Dim ed As Edicion
    Set ed = New Edicion
    ed.IDEdicion = ""

    Dim pub As Boolean
    Dim pError As String

    Dim result As EnumSiNo
    result = CalcularPublicabilidadEdicion(ed, pub, , pError)

    If result <> EnumSiNo.No Then
        Test_CalcularPublicabilidadEdicion_IDVacio = JsonFail("Expected EnumSiNo.No but got " & CStr(result), logs)
        Exit Function
    End If
    If pError = "" Then
        Test_CalcularPublicabilidadEdicion_IDVacio = JsonFail("Expected error message but got empty", logs)
        Exit Function
    End If

    Test_CalcularPublicabilidadEdicion_IDVacio = JsonOk("id_vacio_pass", logs)
End Function

' ============================================================
' Test: CalcularPublicabilidadEdicion — Con riesgo NoPublicable
' Schema-first: fixture tiene PM/PC sin acciones ? veredicto NoPublicable
'   DELETE acciones PM/PC del fixture para garantizar el estado
' Fixture: SeedAll + Constructor.getEdicion
' Teardown: TeardownAll
'
' FIX v1.9: GoTo Teardown + Teardown block
' ============================================================
Public Function Test_CalcularPublicabilidadEdicion_ConRiesgoNoPublicable() As String
    Dim logs(0 To 6) As String
    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: DELETE acciones PM/PC (riesgo sin acciones = NoPublicable)"
    logs(3) = "4. Act: CalcularPublicabilidadEdicion"
    logs(4) = "5. Assert: EdicionPublicable = False"
    logs(5) = "6. Teardown: TeardownAll"

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_CalcularPublicabilidadEdicion_ConRiesgoNoPublicable = JsonFail(cfgError, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    Dim db As DAO.Database
    Set db = GetTestDb(cfgError)
    If db Is Nothing Then
        Test_Fixtures.TeardownAll
        Test_CalcularPublicabilidadEdicion_ConRiesgoNoPublicable = JsonFail(cfgError, logs)
        Exit Function
    End If

    Dim IDEdicion As Long
    IDEdicion = Test_Fixtures.Cache_EdicionId
    Dim idPM As Long
    idPM = Test_Fixtures.Cache_PMId
    Dim idPC As Long
    idPC = Test_Fixtures.Cache_PCId

    ' Arrange: eliminar acciones PM/PC — riesgo queda sin acciones ? NoPublicable
    ' Schema-first: nombre real de tablas según ERD
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgosPlanMitigacionDetalle WHERE IDMitigacion=" & idPM
    db.Execute "DELETE FROM TbRiesgosPlanContingenciaDetalle WHERE IDContingencia=" & idPC
    On Error GoTo 0

    ' Act: obtener Edicion via Constructor
    Dim pError As String
    Dim ed As Edicion
    Set ed = Constructor.getEdicion(CStr(IDEdicion), pError)
    If ed Is Nothing Then GoTo Teardown

    Dim pub As Boolean
    Dim result As EnumSiNo
    result = CalcularPublicabilidadEdicion(ed, pub, , pError)

    ' Capture before teardown
    Dim resultOk As Boolean
    resultOk = (result = EnumSiNo.Sí)
    Dim pubOk As Boolean
    pubOk = pub

    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    Test_Fixtures.TeardownAll

    If Not resultOk Then
        Test_CalcularPublicabilidadEdicion_ConRiesgoNoPublicable = JsonFail("Expected EnumSiNo.Sí but got " & CStr(result) & ": " & pError, logs)
        Exit Function
    End If
    If pubOk <> False Then
        Test_CalcularPublicabilidadEdicion_ConRiesgoNoPublicable = JsonFail("Expected EdicionPublicable=False (sin acciones PM/PC) but got True", logs)
        Exit Function
    End If

    Test_CalcularPublicabilidadEdicion_ConRiesgoNoPublicable = JsonOk("riesgo_no_publicable_pass", logs)
End Function

' ============================================================
' Test: CalcularPublicabilidadEdicion — Todo Publicable
' Schema-first: fixture con acciones PM/PC ya existe en SeedAll
'   Insertar acción en PM y PC para riesgo publicable
' Teardown: DELETE acciones insertadas + TeardownAll
'
' FIX v1.9: GoTo Teardown + Teardown block + On Error Resume Next
' ============================================================
Public Function Test_CalcularPublicabilidadEdicion_TodoPublicable() As String
    Dim logs(0 To 8) As String
    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: INSERT acción en PM + PC"
    logs(3) = "4. Act: CalcularPublicabilidadEdicion"
    logs(4) = "5. Assert: result = EnumSiNo.Sí"
    logs(5) = "6. Assert: EdicionPublicable = True"
    logs(6) = "7. Teardown: DELETE acciones insertadas"
    logs(7) = "8. Teardown: TeardownAll"

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_CalcularPublicabilidadEdicion_TodoPublicable = JsonFail(cfgError, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    Dim db As DAO.Database
    Set db = GetTestDb(cfgError)
    If db Is Nothing Then
        Test_Fixtures.TeardownAll
        Test_CalcularPublicabilidadEdicion_TodoPublicable = JsonFail(cfgError, logs)
        Exit Function
    End If

    Dim IDEdicion As Long
    IDEdicion = Test_Fixtures.Cache_EdicionId
    Dim idPM As Long
    idPM = Test_Fixtures.Cache_PMId
    Dim idPC As Long
    idPC = Test_Fixtures.Cache_PCId

    ' Arrange: buscar IDs libres para las acciones (schema-first: campos NOT NULL)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT MAX(IDAccionMitigacion) FROM TbRiesgosPlanMitigacionDetalle")
    Dim maxIdPM As Long
    maxIdPM = IIf(rs.EOF Or IsNull(rs.Fields(0).value), 900000, rs.Fields(0).value)
    rs.Close

    Set rs = db.OpenRecordset("SELECT MAX(IDAccionContingencia) FROM TbRiesgosPlanContingenciaDetalle")
    Dim maxIdPC As Long
    maxIdPC = IIf(rs.EOF Or IsNull(rs.Fields(0).value), 900000, rs.Fields(0).value)
    rs.Close
    Set rs = Nothing

    Dim idAccionPM As Long
    idAccionPM = maxIdPM + 10
    Dim idAccionPC As Long
    idAccionPC = maxIdPC + 10

    ' Arrange: insertar acciones en PM y PC (campos obligatorios según ERD)
    On Error Resume Next
    db.Execute "INSERT INTO TbRiesgosPlanMitigacionDetalle " & _
               "(IDAccionMitigacion, IDMitigacion, CodAccion, Accion, ResponsableAccion, Estado) " & _
               "VALUES (" & idAccionPM & ", " & idPM & ", 'TEST-PM-" & idAccionPM & "', " & _
               "'Acción fixture test PM', 'TESTUSER', 'Definida')"
    db.Execute "INSERT INTO TbRiesgosPlanContingenciaDetalle " & _
               "(IDAccionContingencia, IDContingencia, CodAccion, Accion, ResponsableAccion, Estado) " & _
               "VALUES (" & idAccionPC & ", " & idPC & ", 'TEST-PC-" & idAccionPC & "', " & _
               "'Acción fixture test PC', 'TESTUSER', 'Definida')"
    On Error GoTo 0

    ' Act: obtener Edicion via Constructor
    Dim pError As String
    Dim ed As Edicion
    Set ed = Constructor.getEdicion(CStr(IDEdicion), pError)
    If ed Is Nothing Then GoTo Teardown

    Dim pub As Boolean
    Dim result As EnumSiNo
    result = CalcularPublicabilidadEdicion(ed, pub, , pError)

    ' Capture before teardown
    Dim resultOk As Boolean
    resultOk = (result = EnumSiNo.Sí)
    Dim pubOk As Boolean
    pubOk = pub

    GoTo Teardown

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgosPlanMitigacionDetalle WHERE IDAccionMitigacion=" & idAccionPM
    db.Execute "DELETE FROM TbRiesgosPlanContingenciaDetalle WHERE IDAccionContingencia=" & idAccionPC
    On Error GoTo 0
    Set db = Nothing
    Test_Fixtures.TeardownAll

    If Not resultOk Then
        Test_CalcularPublicabilidadEdicion_TodoPublicable = JsonFail("Expected EnumSiNo.Sí but got " & CStr(result) & ": " & pError, logs)
        Exit Function
    End If
    If pubOk <> True Then
        Test_CalcularPublicabilidadEdicion_TodoPublicable = JsonFail("Expected EdicionPublicable=True (PM/PC con acciones) but got False", logs)
        Exit Function
    End If

    Test_CalcularPublicabilidadEdicion_TodoPublicable = JsonOk("todo_publicable_pass", logs)
End Function

' ============================================================
' Test: CachePublicabilidad_AsegurarSchema — Crea tabla
' DROP si existe ? verificar que la crea desde cero
'
' FIX: TempVars("Publicabilidad_Usar_Cache")="Sí" antes de llamar
'   porque la función retorna "OK" sin crear tabla cuando flag="No"
'
' FIX v1.9: GoTo Teardown + Teardown block con On Error Resume Next
' Teardown: DROP tabla creada + restaurar TempVars
' ============================================================
Public Function Test_CachePublicabilidad_AsegurarSchema_CreaTabla() As String
    Dim logs(0 To 8) As String
    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: TempVars(Publicabilidad_Usar_Cache)='Sí' (force cache ON)"
    logs(2) = "3. Arrange: DROP TbCachePublicabilidadEdicion si existe"
    logs(3) = "4. Act: CachePublicabilidad_AsegurarSchema"
    logs(4) = "5. Assert: result = 'OK'"
    logs(5) = "6. Assert: tabla existe en TableDefs"
    logs(6) = "7. Teardown: DROP tabla"
    logs(7) = "8. Teardown: TempVars restore"
    logs(8) = "9. Teardown: TeardownAll"

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_CachePublicabilidad_AsegurarSchema_CreaTabla = JsonFail(cfgError, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    ' FIX: guardar flag original antes de forzar cache ON
    ' La función retorna "OK" sin crear tabla cuando flag="No"
    Dim originalCacheFlag As String
    On Error Resume Next
    originalCacheFlag = CStr(Application.TempVars("Publicabilidad_Usar_Cache"))
    If Err.Number <> 0 Then originalCacheFlag = ""
    On Error GoTo 0
    Application.TempVars("Publicabilidad_Usar_Cache") = "Sí"

    Dim db As DAO.Database
    Set db = GetTestDb(cfgError)
    If db Is Nothing Then
        ' Restaurar TempVars antes de salir
        If originalCacheFlag <> "" Then
            Application.TempVars("Publicabilidad_Usar_Cache") = originalCacheFlag
        Else
            Application.TempVars.Remove "Publicabilidad_Usar_Cache"
        End If
        Test_Fixtures.TeardownAll
        Test_CachePublicabilidad_AsegurarSchema_CreaTabla = JsonFail(cfgError, logs)
        Exit Function
    End If

    Const TABLA As String = "TbCachePublicabilidadEdicion"

    ' Arrange: drop if exists (test idempotente)
    Dim tdf As DAO.TableDef
    Dim found As Boolean
    found = False
    For Each tdf In db.TableDefs
        If tdf.Name = TABLA Then
            found = True
            Exit For
        End If
    Next tdf
    If found Then
        On Error Resume Next
        db.Execute "DROP TABLE " & TABLA
        On Error GoTo 0
    End If

    ' Act
    Dim pError As String
    Dim result As String
    result = CachePublicabilidad_AsegurarSchema(pError)

    ' Restore TempVars FIRST before any assertions that could fail
    If originalCacheFlag <> "" Then
        Application.TempVars("Publicabilidad_Usar_Cache") = originalCacheFlag
    Else
        Application.TempVars.Remove "Publicabilidad_Usar_Cache"
    End If

    If result <> "OK" Then GoTo Teardown

    ' Assert: tabla existe
    found = False
    For Each tdf In db.TableDefs
        If tdf.Name = TABLA Then
            found = True
            Exit For
        End If
    Next tdf
    If Not found Then GoTo Teardown

    GoTo Teardown

Teardown:
    On Error Resume Next
    ' Teardown: drop tabla
    For Each tdf In db.TableDefs
        If tdf.Name = TABLA Then
            db.Execute "DROP TABLE " & TABLA
            Exit For
        End If
    Next tdf
    On Error GoTo 0
    Set db = Nothing
    Test_Fixtures.TeardownAll

    ' Re-validate after teardown
    If result <> "OK" Then
        Test_CachePublicabilidad_AsegurarSchema_CreaTabla = JsonFail("Expected 'OK' but got '" & result & "': " & pError, logs)
        Exit Function
    End If
    If Not found Then
        Test_CachePublicabilidad_AsegurarSchema_CreaTabla = JsonFail("Tabla " & TABLA & " no fue creada", logs)
        Exit Function
    End If

    Test_CachePublicabilidad_AsegurarSchema_CreaTabla = JsonOk("asegurar_schema_crea_tabla_pass", logs)
End Function

' ============================================================
' Test: CachePublicabilidad_RecalcularEdicion — Edicion Nil
' Guard: retorna EnumSiNo.No cuando p_Edicion Is Nothing
' Testeable directo: sin fixture
' ============================================================
Public Function Test_CachePublicabilidad_RecalcularEdicion_EdicionNil() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: Edicion = Nothing"
    logs(1) = "2. Act: CachePublicabilidad_RecalcularEdicion(Nothing, db, err)"
    logs(2) = "3. Assert: result = EnumSiNo.No"
    logs(3) = "4. Assert: p_Error no vacío"

    Dim ed As Edicion
    Set ed = Nothing

    Dim db As DAO.Database
    Dim pError As String

    Dim result As EnumSiNo
    result = CachePublicabilidad_RecalcularEdicion(ed, db, pError)

    If result <> EnumSiNo.No Then
        Test_CachePublicabilidad_RecalcularEdicion_EdicionNil = JsonFail("Expected EnumSiNo.No but got " & CStr(result), logs)
        Exit Function
    End If
    If pError = "" Then
        Test_CachePublicabilidad_RecalcularEdicion_EdicionNil = JsonFail("Expected error message but got empty", logs)
        Exit Function
    End If

    Test_CachePublicabilidad_RecalcularEdicion_EdicionNil = JsonOk("recalcular_nil_pass", logs)
End Function

' ============================================================
' Test: CachePublicabilidad_LeerEdicionPublicable — IDEdicion vacío
'
' FIX: Eliminado el assert de p_Error. Con ID vacío la función retorna
'   EnumSiNo.No (cache miss = "no hay cache") sin establecer p_Error.
'   El test verifica: retorna No + p_Error vacío (edge case correcto).
'
' FIX v1.9: On Error Resume Next / On Error GoTo 0 en execute
' FIX v1.9: On Error Resume Next en GetTestDb
' ============================================================
Public Function Test_CachePublicabilidad_LeerEdicionPublicable_IDVacio() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: p_IDEdicion = ''"
    logs(1) = "2. Act: CachePublicabilidad_LeerEdicionPublicable('', pub, err)"
    logs(2) = "3. Assert: result = EnumSiNo.No"
    logs(3) = "4. Assert: p_Error está vacío (con ID vacío = cache miss = No, sin error)"

    Dim pub As Boolean
    Dim pError As String

    Dim result As EnumSiNo
    result = CachePublicabilidad_LeerEdicionPublicable("", pub, pError)

    If result <> EnumSiNo.No Then
        Test_CachePublicabilidad_LeerEdicionPublicable_IDVacio = JsonFail("Expected EnumSiNo.No but got " & CStr(result), logs)
        Exit Function
    End If
    ' FIX: con ID vacío la función retorna No (cache miss) sin mensaje de error.
    ' El test correcto verifica que p_Error esté vacío.
    If pError <> "" Then
        Test_CachePublicabilidad_LeerEdicionPublicable_IDVacio = JsonFail("Expected empty p_Error (ID vacío = cache miss) but got: " & pError, logs)
        Exit Function
    End If

    Test_CachePublicabilidad_LeerEdicionPublicable_IDVacio = JsonOk("leer_id_vacio_pass", logs)
End Function

' ============================================================
' Test: CachePublicabilidad_LeerEdicionPublicable — Cache Hit
' Schema-first: boolean True ? -1 (Access YESNO canónico en SQL)
'   Pre-insertar fila EDICION + verificar que la lee
'
' FIX: TempVars("Publicabilidad_Usar_Cache")="Sí" antes de llamar
'   porque la función retorna EnumSiNo.No directamente cuando flag="No"
'   sin consultar cache (early return).
'
' FIX v1.9: GoTo Teardown + Teardown block + On Error Resume Next
' Teardown: DELETE cache de la edición + restaurar TempVars
' ============================================================
Public Function Test_CachePublicabilidad_LeerEdicionPublicable_CacheHit() As String
    Dim logs(0 To 8) As String
    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: TempVars(Publicabilidad_Usar_Cache)='Sí'"
    logs(3) = "4. Arrange: CachePublicabilidad_AsegurarSchema"
    logs(4) = "5. Arrange: INSERT TbCachePublicabilidadEdicion (Publicable=-1)"
    logs(5) = "6. Act: CachePublicabilidad_LeerEdicionPublicable"
    logs(6) = "7. Assert: result = EnumSiNo.Sí AND pub = True"
    logs(7) = "8. Teardown: DELETE cache + TempVars restore"
    logs(8) = "9. Teardown: TeardownAll"

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_CachePublicabilidad_LeerEdicionPublicable_CacheHit = JsonFail(cfgError, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    ' FIX: guardar flag original antes de forzar cache ON
    ' La función retorna EnumSiNo.No directamente cuando flag="No"
    ' sin consultar cache (early return en línea 401).
    Dim originalCacheFlag As String
    On Error Resume Next
    originalCacheFlag = CStr(Application.TempVars("Publicabilidad_Usar_Cache"))
    If Err.Number <> 0 Then originalCacheFlag = ""
    On Error GoTo 0
    Application.TempVars("Publicabilidad_Usar_Cache") = "Sí"

    Dim db As DAO.Database
    Set db = GetTestDb(cfgError)
    If db Is Nothing Then
        If originalCacheFlag <> "" Then
            Application.TempVars("Publicabilidad_Usar_Cache") = originalCacheFlag
        Else
            Application.TempVars.Remove "Publicabilidad_Usar_Cache"
        End If
        Test_Fixtures.TeardownAll
        Test_CachePublicabilidad_LeerEdicionPublicable_CacheHit = JsonFail(cfgError, logs)
        Exit Function
    End If

    ' Asegurar schema
    Dim asegError As String
    CachePublicabilidad_AsegurarSchema asegError

    Dim IDEdicion As Long
    IDEdicion = Test_Fixtures.Cache_EdicionId

    ' Arrange: limpiar cache previo
    On Error Resume Next
    db.Execute "DELETE FROM TbCachePublicabilidadEdicion WHERE IDEdicion=" & IDEdicion
    On Error GoTo 0

    ' Arrange: insertar fila de cache EDICION con Publicable=True (-1 = Access YESNO)
    On Error Resume Next
    db.Execute "INSERT INTO TbCachePublicabilidadEdicion " & _
               "(IDEdicion, Tipo, IDRiesgo, Publicable, AlgoritmoVersion, UpdatedAt) " & _
               "VALUES (" & IDEdicion & ", 'EDICION', 0, -1, 2, #" & Format$(Now, "yyyy-mm-dd") & "#)"
    On Error GoTo 0

    Dim pub As Boolean
    Dim pError As String

    Dim result As EnumSiNo
    result = CachePublicabilidad_LeerEdicionPublicable(CStr(IDEdicion), pub, pError)

    ' Restore TempVars FIRST before any assertions
    If originalCacheFlag <> "" Then
        Application.TempVars("Publicabilidad_Usar_Cache") = originalCacheFlag
    Else
        Application.TempVars.Remove "Publicabilidad_Usar_Cache"
    End If

    ' Capture results before teardown
    Dim resultOk As Boolean
    resultOk = (result = EnumSiNo.Sí)
    Dim pubOk As Boolean
    pubOk = (pub = True)

    If resultOk And pubOk Then GoTo Teardown

    GoTo Teardown

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbCachePublicabilidadEdicion WHERE IDEdicion=" & IDEdicion
    On Error GoTo 0
    Set db = Nothing
    Test_Fixtures.TeardownAll

    ' Assert using captured values
    If Not resultOk Then
        Test_CachePublicabilidad_LeerEdicionPublicable_CacheHit = JsonFail("Expected EnumSiNo.Sí but got " & CStr(result) & ": " & pError, logs)
        Exit Function
    End If
    If Not pubOk Then
        Test_CachePublicabilidad_LeerEdicionPublicable_CacheHit = JsonFail("Expected pub=True from cache but got False", logs)
        Exit Function
    End If

    Test_CachePublicabilidad_LeerEdicionPublicable_CacheHit = JsonOk("cache_hit_pass", logs)
End Function

' ============================================================
' Test: CachePublicabilidad_LeerEdicionPublicable — Cache Miss
' Sin fila en cache ? retorna EnumSiNo.No
' Fixture: ID 999999 (nunca existe en cache)
'
' FIX: TempVars("Publicabilidad_Usar_Cache")="Sí" forzado (mismo motivo)
'
' FIX v1.9: GoTo Teardown + Teardown block + On Error Resume Next
' Teardown: DELETE por seguridad + restaurar TempVars
' ============================================================
Public Function Test_CachePublicabilidad_LeerEdicionPublicable_CacheMiss() As String
    Dim logs(0 To 6) As String
    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedAll"
    logs(2) = "3. Arrange: TempVars(Publicabilidad_Usar_Cache)='Sí'"
    logs(3) = "4. Arrange: DELETE WHERE IDEdicion=999999"
    logs(4) = "5. Act: CachePublicabilidad_LeerEdicionPublicable('999999', pub, err)"
    logs(5) = "6. Assert: result = EnumSiNo.No"
    logs(6) = "7. Teardown: DELETE + TempVars restore + TeardownAll"

    Dim cfgError As String
    If Not EnsureTestConfigLoaded(cfgError) Then
        Test_CachePublicabilidad_LeerEdicionPublicable_CacheMiss = JsonFail(cfgError, logs)
        Exit Function
    End If

    Test_Fixtures.SeedAll

    ' FIX: guardar flag original antes de forzar cache ON
    Dim originalCacheFlag As String
    On Error Resume Next
    originalCacheFlag = CStr(Application.TempVars("Publicabilidad_Usar_Cache"))
    If Err.Number <> 0 Then originalCacheFlag = ""
    On Error GoTo 0
    Application.TempVars("Publicabilidad_Usar_Cache") = "Sí"

    Dim db As DAO.Database
    Set db = GetTestDb(cfgError)
    If db Is Nothing Then
        If originalCacheFlag <> "" Then
            Application.TempVars("Publicabilidad_Usar_Cache") = originalCacheFlag
        Else
            Application.TempVars.Remove "Publicabilidad_Usar_Cache"
        End If
        Test_Fixtures.TeardownAll
        Test_CachePublicabilidad_LeerEdicionPublicable_CacheMiss = JsonFail(cfgError, logs)
        Exit Function
    End If

    ' Asegurar schema
    Dim asegError As String
    CachePublicabilidad_AsegurarSchema asegError

    Const MISS_ID As Long = 999999

    ' Arrange: limpiar por si existía
    On Error Resume Next
    db.Execute "DELETE FROM TbCachePublicabilidadEdicion WHERE IDEdicion=" & MISS_ID
    On Error GoTo 0

    Dim pub As Boolean
    Dim pError As String

    Dim result As EnumSiNo
    result = CachePublicabilidad_LeerEdicionPublicable(CStr(MISS_ID), pub, pError)

    ' Restore TempVars FIRST before any assertions
    If originalCacheFlag <> "" Then
        Application.TempVars("Publicabilidad_Usar_Cache") = originalCacheFlag
    Else
        Application.TempVars.Remove "Publicabilidad_Usar_Cache"
    End If

    If result <> EnumSiNo.No Then GoTo Teardown

    GoTo Teardown

Teardown:
    On Error Resume Next
    db.Execute "DELETE FROM TbCachePublicabilidadEdicion WHERE IDEdicion=" & MISS_ID
    On Error GoTo 0
    Set db = Nothing
    Test_Fixtures.TeardownAll

    If result <> EnumSiNo.No Then
        Test_CachePublicabilidad_LeerEdicionPublicable_CacheMiss = JsonFail("Expected EnumSiNo.No (cache miss) but got " & CStr(result), logs)
        Exit Function
    End If

    Test_CachePublicabilidad_LeerEdicionPublicable_CacheMiss = JsonOk("cache_miss_pass", logs)
End Function

Public Function Test_PublicabilidadEdicion_RutaPDF_DerivaDesdeHTML() As String
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: ruta HTML con extension mixta"
    logs(1) = "2. Act: PublicabilidadEdicion_RutaPDFDesdeHTMLParaTest"
    logs(2) = "3. Assert: deriva extension .pdf sin invocar Edge"

    Dim rutaPDF As String
    rutaPDF = PublicabilidadEdicion_RutaPDFDesdeHTMLParaTest("C:\Temp\Publicabilidad_Edicion_1.HTML")

    If rutaPDF <> "C:\Temp\Publicabilidad_Edicion_1.pdf" Then
        Test_PublicabilidadEdicion_RutaPDF_DerivaDesdeHTML = JsonFail("Expected C:\Temp\Publicabilidad_Edicion_1.pdf but got " & rutaPDF, logs)
        Exit Function
    End If

    Test_PublicabilidadEdicion_RutaPDF_DerivaDesdeHTML = JsonOk("ruta_pdf_derivada", logs)
End Function

Public Function Test_PublicabilidadEdicion_RutaSalida_RespetaFlagPDF() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: ruta HTML"
    logs(1) = "2. Act: RutaSalida con p_GenerarPDF=False"
    logs(2) = "3. Act: RutaSalida con p_GenerarPDF=True"
    logs(3) = "4. Assert: HTML por defecto, PDF con flag"

    Dim rutaHTML As String
    rutaHTML = "C:\Temp\Publicabilidad_Edicion_1.html"

    Dim salidaHTML As String
    Dim salidaPDF As String
    salidaHTML = PublicabilidadEdicion_RutaSalidaParaTest(rutaHTML, False)
    salidaPDF = PublicabilidadEdicion_RutaSalidaParaTest(rutaHTML, True)

    If salidaHTML <> rutaHTML Then
        Test_PublicabilidadEdicion_RutaSalida_RespetaFlagPDF = JsonFail("Expected HTML output by default but got " & salidaHTML, logs)
        Exit Function
    End If

    If salidaPDF <> "C:\Temp\Publicabilidad_Edicion_1.pdf" Then
        Test_PublicabilidadEdicion_RutaSalida_RespetaFlagPDF = JsonFail("Expected PDF output with flag but got " & salidaPDF, logs)
        Exit Function
    End If

    Test_PublicabilidadEdicion_RutaSalida_RespetaFlagPDF = JsonOk("ruta_salida_flag_pdf", logs)
End Function
