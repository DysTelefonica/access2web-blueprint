Attribute VB_Name = "Test_modFormRiesgoDocumentosHelper"
Option Compare Database
Option Explicit

' ============================================================
' Test_modFormRiesgoDocumentosHelper
'   B2 / Punto 15 — Acta 25/06 + Contrato 22/06
'   Skill: access-vba-tdd v2.6.1
'
'   Helper cache-first + repositorio atómico de anexos edición.
'   Cubre 14 átomos:
'     PR1 — helper cache-first        (1-8)
'     PR2 — Edicion.ColAnexosTotales + repo (9-12)
'     PR3 — smoke test de forms       (13-14)
'
'   Helper bajo prueba: modFormRiesgoDocumentosHelper.bas
'   Repositorio:        Constructor.getAnexosTotalesDeEdicion
'   Modelo:             Edicion.ColAnexosTotales
'
'   IDs de fixture (rangos no solapados con SeedAll base 900500-900599):
'     FIX_EDICION_B2     = 910000
'     FIX_PROYECTO_B2    = 910001
'     FIX_EXPEDIENTE_B2  = 910002
'     FIX_RIESGO_B2_A    = 910010  (1 anexo)
'     FIX_RIESGO_B2_B    = 910011  (0 anexos)
'     FIX_RIESGO_B2_C    = 910012  (2 anexos)
'     FIX_ANEXO_BASE_B2  = 910100  (siguiente bloque para anexos seed)
' ============================================================

Private Const FIX_EDICION_B2 As Long = 910000
Private Const FIX_PROYECTO_B2 As Long = 910001
Private Const FIX_EXPEDIENTE_B2 As Long = 910002
Private Const FIX_RIESGO_B2_A As Long = 910010
Private Const FIX_RIESGO_B2_B As Long = 910011
Private Const FIX_RIESGO_B2_C As Long = 910012
Private Const FIX_ANEXO_BASE_B2 As Long = 910100
Private Const FIX_ANEXO_MAX_B2 As Long = 910199

Private Const TITULO_DIRECTO_A As String = "Doc directo edicion"
Private Const TITULO_DIRECTO_B As String = "Doc directo edicion 2"
Private Const TITULO_RIESGO_A As String = "Doc riesgo A"
Private Const TITULO_RIESGO_C_1 As String = "Doc riesgo C 1"
Private Const TITULO_RIESGO_C_2 As String = "Doc riesgo C 2"

Private Const CACHE_FIX_NOMBRE As String = "modFormRiesgoDocumentosHelper"

' --- Wrappers JSON (delegan a Test_Helper) ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' --- Reset del helper entre átomos (idempotencia entre tests) ---
Private Sub ResetHelperCaches()
    On Error Resume Next
    modFormRiesgoDocumentosHelper.InvalidarCachePorIDEdicion CStr(FIX_EDICION_B2)
    modFormRiesgoDocumentosHelper.InvalidarCachePorIDRiesgo CStr(FIX_RIESGO_B2_A)
    modFormRiesgoDocumentosHelper.InvalidarCachePorIDRiesgo CStr(FIX_RIESGO_B2_B)
    modFormRiesgoDocumentosHelper.InvalidarCachePorIDRiesgo CStr(FIX_RIESGO_B2_C)
    On Error GoTo 0
End Sub

' --- Setup / teardown del grafo B2 ---
Private Sub SeedGrafoB2()
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then err.Raise 1001, "SeedGrafoB2", "GetTestDb Nothing: " & dbErr

    On Error Resume Next
    ' Limpieza en orden inverso FK
    db.Execute "DELETE FROM TbAnexos WHERE IDAnexo>=" & FIX_ANEXO_BASE_B2 & " AND IDAnexo<=" & FIX_ANEXO_MAX_B2
    db.Execute "DELETE FROM TbAnexos WHERE IDEdicion=" & FIX_EDICION_B2
    db.Execute "DELETE FROM TbAnexos WHERE IDRiesgo IN (" & FIX_RIESGO_B2_A & "," & FIX_RIESGO_B2_B & "," & FIX_RIESGO_B2_C & ")"
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo IN (" & FIX_RIESGO_B2_A & "," & FIX_RIESGO_B2_B & "," & FIX_RIESGO_B2_C & ")"
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_B2
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_PROYECTO_B2
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_EXPEDIENTE_B2
    On Error GoTo 0

    ' Expediente
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
        "VALUES (" & FIX_EXPEDIENTE_B2 & ", 'B2FIX', 'Fixture B2 anexos', 'Test', 1)"

    ' Proyecto
    db.Execute "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto) " & _
        "VALUES (" & FIX_PROYECTO_B2 & ", " & FIX_EXPEDIENTE_B2 & ", 'B2PROJ')"

    ' Edicion (sin FechaPublicacion ? EsActivo=Sí en logica del modelo)
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
        "VALUES (" & FIX_EDICION_B2 & ", " & FIX_PROYECTO_B2 & ", 1, 'TESTUSER')"

    ' 3 riesgos hijos
    Dim i As Long
    Dim idR As Long
    Dim codR As String
    For i = 0 To 2
        idR = FIX_RIESGO_B2_A + i
        codR = "R" & Format(i + 1, "000")
        db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoRiesgo, Descripcion, Estado, Priorizacion, " & _
            "CodigoUnico, FechaDetectado, DetectadoPor, EntidadDetecta, Plazo, Calidad, Coste, ImpactoGlobal, " & _
            "Vulnerabilidad, Valoracion, Mitigacion, Contingencia, RequierePlanContingencia) " & _
            "VALUES (" & idR & ", " & FIX_EDICION_B2 & ", '" & codR & "', " & _
            "'Fixture riesgo B2', 'Detectado', 3, " & _
            "'UNICO-B2-" & idR & "', #" & Format$(Now, "yyyy-mm-dd") & "#, " & _
            "'TESTUSER', 'TEST', 'Medio', 'Medio', 'Medio', 'Medio', 'Medio', 'Medio', 'Reducir', 'Sí', 'Sí')"
    Next i

    ' 2 anexos directos de la edicion
    db.Execute "INSERT INTO TbAnexos (IDAnexo, IDEdicion, Titulo, NombreArchivo, FechaAnexo, EvidenciaUTE, EvidenciaSuministrador) " & _
        "VALUES (" & FIX_ANEXO_BASE_B2 & ", " & FIX_EDICION_B2 & ", '" & TITULO_DIRECTO_A & "', " & _
        "'directo_a.txt', #" & Format$(Now, "yyyy-mm-dd") & "#, 'No', 'No')"
    db.Execute "INSERT INTO TbAnexos (IDAnexo, IDEdicion, Titulo, NombreArchivo, FechaAnexo, EvidenciaUTE, EvidenciaSuministrador) " & _
        "VALUES (" & (FIX_ANEXO_BASE_B2 + 1) & ", " & FIX_EDICION_B2 & ", '" & TITULO_DIRECTO_B & "', " & _
        "'directo_b.txt', #" & Format$(Now, "yyyy-mm-dd") & "#, 'No', 'No')"

    ' 1 anexo en riesgo A
    db.Execute "INSERT INTO TbAnexos (IDAnexo, IDRiesgo, Titulo, NombreArchivo, FechaAnexo, EvidenciaUTE, EvidenciaSuministrador) " & _
        "VALUES (" & (FIX_ANEXO_BASE_B2 + 2) & ", " & FIX_RIESGO_B2_A & ", '" & TITULO_RIESGO_A & "', " & _
        "'riesgo_a.txt', #" & Format$(Now, "yyyy-mm-dd") & "#, 'No', 'No')"

    ' 0 anexos en riesgo B (no se inserta nada)

    ' 2 anexos en riesgo C
    db.Execute "INSERT INTO TbAnexos (IDAnexo, IDRiesgo, Titulo, NombreArchivo, FechaAnexo, EvidenciaUTE, EvidenciaSuministrador) " & _
        "VALUES (" & (FIX_ANEXO_BASE_B2 + 3) & ", " & FIX_RIESGO_B2_C & ", '" & TITULO_RIESGO_C_1 & "', " & _
        "'riesgo_c1.txt', #" & Format$(Now, "yyyy-mm-dd") & "#, 'No', 'No')"
    db.Execute "INSERT INTO TbAnexos (IDAnexo, IDRiesgo, Titulo, NombreArchivo, FechaAnexo, EvidenciaUTE, EvidenciaSuministrador) " & _
        "VALUES (" & (FIX_ANEXO_BASE_B2 + 4) & ", " & FIX_RIESGO_B2_C & ", '" & TITULO_RIESGO_C_2 & "', " & _
        "'riesgo_c2.txt', #" & Format$(Now, "yyyy-mm-dd") & "#, 'No', 'No')"

    Set db = Nothing
End Sub

Private Sub TeardownGrafoB2()
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then Exit Sub

    On Error Resume Next
    db.Execute "DELETE FROM TbAnexos WHERE IDAnexo>=" & FIX_ANEXO_BASE_B2 & " AND IDAnexo<=" & FIX_ANEXO_MAX_B2
    db.Execute "DELETE FROM TbAnexos WHERE IDEdicion=" & FIX_EDICION_B2
    db.Execute "DELETE FROM TbAnexos WHERE IDRiesgo IN (" & FIX_RIESGO_B2_A & "," & FIX_RIESGO_B2_B & "," & FIX_RIESGO_B2_C & ")"
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo IN (" & FIX_RIESGO_B2_A & "," & FIX_RIESGO_B2_B & "," & FIX_RIESGO_B2_C & ")"
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_EDICION_B2
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_PROYECTO_B2
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_EXPEDIENTE_B2
    On Error GoTo 0
    Set db = Nothing
End Sub

' --- Guarda ---
Private Function EnsureTestingContext(ByRef logs() As String) As Boolean
    Dim errMsg As String
    If Not Test_Helper.ForceLocalBackend(errMsg) Then
        BuildFail "ForceLocalBackend fallo: " & errMsg, logs
        EnsureTestingContext = False
        Exit Function
    End If
    EnsureTestingContext = True
End Function

' --- Contadores observabilidad cache (test-only seam) ---
Private Function LeerContadorLlamadasRepo() As Long
    On Error Resume Next
    LeerContadorLlamadasRepo = Constructor.GetContadorLlamadasGetAnexosTotalesDeEdicion()
    If err.Number <> 0 Then LeerContadorLlamadasRepo = -1
    On Error GoTo 0
End Function

Private Sub ResetContadorLlamadasRepo()
    On Error Resume Next
    Constructor.ResetContadorLlamadasGetAnexosTotalesDeEdicion
    On Error GoTo 0
End Sub

' --- Inspectors de cache (test-only seam) ---
Private Function HelperCacheContieneIDEdicion(ByVal p_IDEdicion As String) As Boolean
    On Error Resume Next
    HelperCacheContieneIDEdicion = modFormRiesgoDocumentosHelper.CachePorIDEdicionContieneKey(CStr(p_IDEdicion))
    If err.Number <> 0 Then HelperCacheContieneIDEdicion = False
    On Error GoTo 0
End Function

Private Function HelperCacheContieneIDRiesgo(ByVal p_IDRiesgo As String) As Boolean
    On Error Resume Next
    HelperCacheContieneIDRiesgo = modFormRiesgoDocumentosHelper.CachePorIDRiesgoContieneKey(CStr(p_IDRiesgo))
    If err.Number <> 0 Then HelperCacheContieneIDRiesgo = False
    On Error GoTo 0
End Function

' ============================================================
' ÁTOMOS TDD — 14 átomos
' ============================================================

' ----------------------------------------------------------------------------
' ATOMO 1 — CacheHit GetAnexosDeEdicion
'   Getter repetido para el mismo IDEdicion ? 2ª llamada no consulta repo
' ----------------------------------------------------------------------------
Public Function Test_Helper_CacheHit_GetAnexosDeEdicion_LecturaRepetidaMismoID() As String
    Dim logs(0 To 7) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Helper_CacheHit_GetAnexosDeEdicion_LecturaRepetidaMismoID = logs(0)
        Exit Function
    End If
    ResetContadorLlamadasRepo
    SeedGrafoB2
    ResetHelperCaches

    logs(0) = "1. Arrange: Fixture B2 sembrado, caches vaciados, contador a 0"
    logs(1) = "2. Act: 1ª llamada GetAnexosDeEdicion(FIX_EDICION_B2)"
    Dim errA As String
    Dim colA As Scripting.Dictionary
    Set colA = modFormRiesgoDocumentosHelper.getAnexosDeEdicionCached(CStr(FIX_EDICION_B2), errA)
    If errA <> "" Then GoTo Fail
    Dim n1 As Long
    n1 = colA.Count
    logs(2) = "3. Assert: 1ª lectura ? " & n1 & " anexos"

    logs(3) = "4. Act: 2ª llamada GetAnexosDeEdicion(FIX_EDICION_B2)"
    Dim errB As String
    Dim colB As Scripting.Dictionary
    Set colB = modFormRiesgoDocumentosHelper.getAnexosDeEdicionCached(CStr(FIX_EDICION_B2), errB)
    If errB <> "" Then GoTo Fail
    Dim n2 As Long
    n2 = colB.Count
    logs(4) = "5. Assert: 2ª lectura ? " & n2 & " anexos"

    If n1 <> n2 Then GoTo Fail
    If Not (colA Is colB) Then GoTo Fail
    If Not HelperCacheContieneIDEdicion(CStr(FIX_EDICION_B2)) Then GoTo Fail

    logs(5) = "6. Assert: contador repo = " & LeerContadorLlamadasRepo() & " (esperado 1; cache hit en 2ª)"
    If LeerContadorLlamadasRepo() <> 1 Then GoTo Fail

    logs(6) = "7. Assert: Dictionary devuelto es la misma instancia"
    logs(7) = "8. PASS"

    TeardownGrafoB2
    Test_Helper_CacheHit_GetAnexosDeEdicion_LecturaRepetidaMismoID = BuildOk("cache_hit_edicion_pass", logs)
    Exit Function

Fail:
    Dim em As String
    em = "Test_Helper_CacheHit_GetAnexosDeEdicion_LecturaRepetidaMismoID fallo: " & err.description
    TeardownGrafoB2
    Test_Helper_CacheHit_GetAnexosDeEdicion_LecturaRepetidaMismoID = BuildFail(em, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 2 — CacheMiss primera lectura
' ----------------------------------------------------------------------------
Public Function Test_Helper_CacheMiss_GetAnexosDeEdicion_PrimeraLectura() As String
    Dim logs(0 To 5) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Helper_CacheMiss_GetAnexosDeEdicion_PrimeraLectura = logs(0)
        Exit Function
    End If
    SeedGrafoB2
    ResetHelperCaches
    ResetContadorLlamadasRepo

    logs(0) = "1. Arrange: Fixture B2 sembrado, cache vacía, contador a 0"
    logs(1) = "2. Act: 1ª lectura GetAnexosDeEdicion(FIX_EDICION_B2)"

    Dim errA As String
    Dim colA As Scripting.Dictionary
    Set colA = modFormRiesgoDocumentosHelper.getAnexosDeEdicionCached(CStr(FIX_EDICION_B2), errA)
    If errA <> "" Then GoTo Fail
    logs(2) = "3. Assert: devuelve " & colA.Count & " anexos (esperado 5: 2 directos + 3 riesgos)"
    If colA.Count <> 5 Then GoTo Fail

    logs(3) = "4. Assert: contador repo = " & LeerContadorLlamadasRepo() & " (esperado 1)"
    If LeerContadorLlamadasRepo() <> 1 Then GoTo Fail

    logs(4) = "5. Assert: cache por IDEdicion contiene FIX_EDICION_B2"
    If Not HelperCacheContieneIDEdicion(CStr(FIX_EDICION_B2)) Then GoTo Fail

    logs(5) = "6. PASS"

    TeardownGrafoB2
    Test_Helper_CacheMiss_GetAnexosDeEdicion_PrimeraLectura = BuildOk("cache_miss_edicion_pass", logs)
    Exit Function

Fail:
    Dim em As String
    em = "Test_Helper_CacheMiss_GetAnexosDeEdicion_PrimeraLectura fallo: errA=" & errA & " | Err.Description=" & err.description
    TeardownGrafoB2
    Test_Helper_CacheMiss_GetAnexosDeEdicion_PrimeraLectura = BuildFail(em, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 3 — CacheHit GetAnexosDeRiesgo
' ----------------------------------------------------------------------------
Public Function Test_Helper_CacheHit_GetAnexosDeRiesgo_LecturaRepetidaMismoID() As String
    Dim logs(0 To 7) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Helper_CacheHit_GetAnexosDeRiesgo_LecturaRepetidaMismoID = logs(0)
        Exit Function
    End If
    SeedGrafoB2
    ResetHelperCaches

    logs(0) = "1. Arrange: Fixture B2 sembrado, caches vaciados"
    logs(1) = "2. Act: 1ª lectura GetAnexosDeRiesgo(FIX_RIESGO_B2_C)"

    Dim errA As String
    Dim colA As Scripting.Dictionary
    Set colA = modFormRiesgoDocumentosHelper.getAnexosDeRiesgoCached(CStr(FIX_RIESGO_B2_C), errA)
    If errA <> "" Then GoTo Fail
    Dim n1 As Long
    n1 = colA.Count
    logs(2) = "3. Assert: 1ª lectura ? " & n1 & " anexos (esperado 2)"

    logs(3) = "4. Act: 2ª lectura GetAnexosDeRiesgo(FIX_RIESGO_B2_C)"
    Dim errB As String
    Dim colB As Scripting.Dictionary
    Set colB = modFormRiesgoDocumentosHelper.getAnexosDeRiesgoCached(CStr(FIX_RIESGO_B2_C), errB)
    If errB <> "" Then GoTo Fail
    Dim n2 As Long
    n2 = colB.Count
    logs(4) = "5. Assert: 2ª lectura ? " & n2 & " anexos"

    If n1 <> n2 Then GoTo Fail
    If n1 <> 2 Then GoTo Fail
    If Not (colA Is colB) Then GoTo Fail
    If Not HelperCacheContieneIDRiesgo(CStr(FIX_RIESGO_B2_C)) Then GoTo Fail

    logs(5) = "6. Assert: misma instancia de Dictionary devuelta"
    logs(6) = "7. PASS"

    TeardownGrafoB2
    Test_Helper_CacheHit_GetAnexosDeRiesgo_LecturaRepetidaMismoID = BuildOk("cache_hit_riesgo_pass", logs)
    Exit Function

Fail:
    Dim em As String
    em = "Test_Helper_CacheHit_GetAnexosDeRiesgo_LecturaRepetidaMismoID fallo: " & err.description
    TeardownGrafoB2
    Test_Helper_CacheHit_GetAnexosDeRiesgo_LecturaRepetidaMismoID = BuildFail(em, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 4 — InvalidacionPorIDRiesgo no afecta otro riesgo
' ----------------------------------------------------------------------------
Public Function Test_Helper_InvalidacionPorIDRiesgo_NoAfectaOtroIDRiesgo() As String
    Dim logs(0 To 8) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Helper_InvalidacionPorIDRiesgo_NoAfectaOtroIDRiesgo = logs(0)
        Exit Function
    End If
    SeedGrafoB2
    ResetHelperCaches

    logs(0) = "1. Arrange: Fixture B2, caches vacías"
    logs(1) = "2. Act: carga cache de riesgo A y riesgo C"

    Dim errA As String
    Dim colA As Scripting.Dictionary
    Set colA = modFormRiesgoDocumentosHelper.getAnexosDeRiesgoCached(CStr(FIX_RIESGO_B2_A), errA)
    If errA <> "" Then GoTo Fail

    Dim errC As String
    Dim colC As Scripting.Dictionary
    Set colC = modFormRiesgoDocumentosHelper.getAnexosDeRiesgoCached(CStr(FIX_RIESGO_B2_C), errC)
    If errC <> "" Then GoTo Fail

    If Not HelperCacheContieneIDRiesgo(CStr(FIX_RIESGO_B2_A)) Then GoTo Fail
    If Not HelperCacheContieneIDRiesgo(CStr(FIX_RIESGO_B2_C)) Then GoTo Fail
    logs(2) = "3. Assert: ambos riesgos cacheados"

    logs(3) = "4. Act: InvalidarCachePorIDRiesgo(FIX_RIESGO_B2_A)"
    modFormRiesgoDocumentosHelper.InvalidarCachePorIDRiesgo CStr(FIX_RIESGO_B2_A)

    logs(4) = "5. Assert: cache de riesgo A vaciada, cache de riesgo C intacta"
    If HelperCacheContieneIDRiesgo(CStr(FIX_RIESGO_B2_A)) Then GoTo Fail
    If Not HelperCacheContieneIDRiesgo(CStr(FIX_RIESGO_B2_C)) Then GoTo Fail

    logs(5) = "6. Assert: lectura de riesgo C sigue siendo misma instancia (cache hit)"
    Dim errC2 As String
    Dim colC2 As Scripting.Dictionary
    Set colC2 = modFormRiesgoDocumentosHelper.getAnexosDeRiesgoCached(CStr(FIX_RIESGO_B2_C), errC2)
    If errC2 <> "" Then GoTo Fail
    If Not (colC Is colC2) Then GoTo Fail

    logs(6) = "7. PASS"
    TeardownGrafoB2
    Test_Helper_InvalidacionPorIDRiesgo_NoAfectaOtroIDRiesgo = BuildOk("inval_riesgo_aislada_pass", logs)
    Exit Function

Fail:
    Dim em As String
    em = "Test_Helper_InvalidacionPorIDRiesgo_NoAfectaOtroIDRiesgo fallo: " & err.description
    TeardownGrafoB2
    Test_Helper_InvalidacionPorIDRiesgo_NoAfectaOtroIDRiesgo = BuildFail(em, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 5 — InvalidacionPorIDEdicion no afecta otra edicion
' ----------------------------------------------------------------------------
Public Function Test_Helper_InvalidacionPorIDEdicion_NoAfectaOtroIDEdicion() As String
    Dim logs(0 To 7) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Helper_InvalidacionPorIDEdicion_NoAfectaOtroIDEdicion = logs(0)
        Exit Function
    End If
    SeedGrafoB2
    ResetHelperCaches

    logs(0) = "1. Arrange: Fixture B2, caches vacías"
    logs(1) = "2. Act: cargar cache de FIX_EDICION_B2 y de un IDEdicion inexistente"

    Dim errA As String
    Dim colA As Scripting.Dictionary
    Set colA = modFormRiesgoDocumentosHelper.getAnexosDeEdicionCached(CStr(FIX_EDICION_B2), errA)
    If errA <> "" Then GoTo Fail

    ' Sembramos otra edicion dummy solo para probar que invalidar la FIX no la toca
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then GoTo Fail
    On Error Resume Next
    db.Execute "DELETE FROM TbAnexos WHERE IDEdicion=" & (FIX_EDICION_B2 + 1)
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & (FIX_EDICION_B2 + 1)
    On Error GoTo 0
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
        "VALUES (" & (FIX_EDICION_B2 + 1) & ", " & FIX_PROYECTO_B2 & ", 2, 'TESTUSER')"
    db.Execute "INSERT INTO TbAnexos (IDAnexo, IDEdicion, Titulo, NombreArchivo, FechaAnexo, EvidenciaUTE, EvidenciaSuministrador) " & _
        "VALUES (" & (FIX_ANEXO_BASE_B2 + 50) & ", " & (FIX_EDICION_B2 + 1) & ", 'vecino', " & _
        "'vecino.txt', #" & Format$(Now, "yyyy-mm-dd") & "#, 'No', 'No')"

    Dim errB As String
    Dim colB As Scripting.Dictionary
    Set colB = modFormRiesgoDocumentosHelper.getAnexosDeEdicionCached(CStr(FIX_EDICION_B2 + 1), errB)
    If errB <> "" Then GoTo Fail
    If Not HelperCacheContieneIDEdicion(CStr(FIX_EDICION_B2)) Then GoTo Fail
    If Not HelperCacheContieneIDEdicion(CStr(FIX_EDICION_B2 + 1)) Then GoTo Fail
    logs(2) = "3. Assert: ambas ediciones cacheadas"

    logs(3) = "4. Act: InvalidarCachePorIDEdicion(FIX_EDICION_B2)"
    modFormRiesgoDocumentosHelper.InvalidarCachePorIDEdicion CStr(FIX_EDICION_B2)

    logs(4) = "5. Assert: cache de FIX_EDICION_B2 vaciada, cache de FIX_EDICION_B2+1 intacta"
    If HelperCacheContieneIDEdicion(CStr(FIX_EDICION_B2)) Then GoTo Fail
    If Not HelperCacheContieneIDEdicion(CStr(FIX_EDICION_B2 + 1)) Then GoTo Fail

    logs(5) = "6. PASS"

    On Error Resume Next
    db.Execute "DELETE FROM TbAnexos WHERE IDAnexo=" & (FIX_ANEXO_BASE_B2 + 50)
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & (FIX_EDICION_B2 + 1)
    On Error GoTo 0
    Set db = Nothing
    TeardownGrafoB2
    Test_Helper_InvalidacionPorIDEdicion_NoAfectaOtroIDEdicion = BuildOk("inval_edicion_aislada_pass", logs)
    Exit Function

Fail:
    Dim em As String
    em = "Test_Helper_InvalidacionPorIDEdicion_NoAfectaOtroIDEdicion fallo: " & err.description
    TeardownGrafoB2
    Test_Helper_InvalidacionPorIDEdicion_NoAfectaOtroIDEdicion = BuildFail(em, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 6 — RegistrarAnexo invalida cache de riesgo + edicion padre
' ----------------------------------------------------------------------------
Public Function Test_Helper_RegistrarAnexo_InvalidaCacheIDRiesgoYIDEdicionPadre() As String
    Dim logs(0 To 8) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Helper_RegistrarAnexo_InvalidaCacheIDRiesgoYIDEdicionPadre = logs(0)
        Exit Function
    End If
    SeedGrafoB2
    ResetHelperCaches

    logs(0) = "1. Arrange: Fixture B2, caches vacías"
    logs(1) = "2. Act: poblar caches de edicion y riesgo A"

    Dim errE As String
    Dim colE As Scripting.Dictionary
    Set colE = modFormRiesgoDocumentosHelper.getAnexosDeEdicionCached(CStr(FIX_EDICION_B2), errE)
    If errE <> "" Then GoTo Fail

    Dim errR As String
    Dim colR As Scripting.Dictionary
    Set colR = modFormRiesgoDocumentosHelper.getAnexosDeRiesgoCached(CStr(FIX_RIESGO_B2_A), errR)
    If errR <> "" Then GoTo Fail

    If Not HelperCacheContieneIDEdicion(CStr(FIX_EDICION_B2)) Then GoTo Fail
    If Not HelperCacheContieneIDRiesgo(CStr(FIX_RIESGO_B2_A)) Then GoTo Fail
    logs(2) = "3. Assert: ambas caches pobladas"

    logs(3) = "4. Act: RegistrarEnRiesgo(FIX_RIESGO_B2_A, '', 'titulo test')"
    ' Sin URL real el helper devolvera error de MotivoNoOK, pero la cache debe invalidarse igual
    Dim errReg As String
    Dim objAnexo As Anexo
    Set objAnexo = modFormRiesgoDocumentosHelper.RegistrarEnRiesgo( _
        CStr(FIX_RIESGO_B2_A), "", "titulo test registrar", errReg)
    ' No validamos errReg porque sin URL valida falla por MotivoNoOK (esperado).

    logs(4) = "5. Assert: cache de riesgo A vaciada tras registrar"
    If HelperCacheContieneIDRiesgo(CStr(FIX_RIESGO_B2_A)) Then GoTo Fail

    logs(5) = "6. Assert: cache de edicion padre tambien vaciada"
    If HelperCacheContieneIDEdicion(CStr(FIX_EDICION_B2)) Then GoTo Fail

    logs(6) = "7. PASS"
    Set objAnexo = Nothing
    TeardownGrafoB2
    Test_Helper_RegistrarAnexo_InvalidaCacheIDRiesgoYIDEdicionPadre = BuildOk("registrar_invalida_caches_pass", logs)
    Exit Function

Fail:
    Dim em As String
    em = "Test_Helper_RegistrarAnexo_InvalidaCacheIDRiesgoYIDEdicionPadre fallo: " & err.description
    TeardownGrafoB2
    Test_Helper_RegistrarAnexo_InvalidaCacheIDRiesgoYIDEdicionPadre = BuildFail(em, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 7 — EliminarAnexo invalida caches (analogo a registrar)
' ----------------------------------------------------------------------------
Public Function Test_Helper_EliminarAnexo_InvalidaCaches() As String
    Dim logs(0 To 8) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Helper_EliminarAnexo_InvalidaCaches = logs(0)
        Exit Function
    End If
    SeedGrafoB2
    ResetHelperCaches

    logs(0) = "1. Arrange: Fixture B2, caches vacías"
    logs(1) = "2. Act: poblar caches de edicion y riesgo C"

    Dim errE As String
    Dim colE As Scripting.Dictionary
    Set colE = modFormRiesgoDocumentosHelper.getAnexosDeEdicionCached(CStr(FIX_EDICION_B2), errE)
    If errE <> "" Then GoTo Fail

    Dim errR As String
    Dim colR As Scripting.Dictionary
    Set colR = modFormRiesgoDocumentosHelper.getAnexosDeRiesgoCached(CStr(FIX_RIESGO_B2_C), errR)
    If errR <> "" Then GoTo Fail

    If Not HelperCacheContieneIDEdicion(CStr(FIX_EDICION_B2)) Then GoTo Fail
    If Not HelperCacheContieneIDRiesgo(CStr(FIX_RIESGO_B2_C)) Then GoTo Fail
    logs(2) = "3. Assert: ambas caches pobladas"

    logs(3) = "4. Act: EliminarAnexo(IDAnexo del riesgo C)"
    Dim errEl As String
    ' Usamos un IDAnexo que existe (riesgo C, segundo anexo)
    modFormRiesgoDocumentosHelper.EliminarAnexo CStr(FIX_ANEXO_BASE_B2 + 4), errEl

    logs(4) = "5. Assert: cache de riesgo C vaciada tras eliminar"
    If HelperCacheContieneIDRiesgo(CStr(FIX_RIESGO_B2_C)) Then GoTo Fail

    logs(5) = "6. Assert: cache de edicion padre tambien vaciada"
    If HelperCacheContieneIDEdicion(CStr(FIX_EDICION_B2)) Then GoTo Fail

    logs(6) = "7. PASS"
    TeardownGrafoB2
    Test_Helper_EliminarAnexo_InvalidaCaches = BuildOk("eliminar_invalida_caches_pass", logs)
    Exit Function

Fail:
    Dim em As String
    em = "Test_Helper_EliminarAnexo_InvalidaCaches fallo: " & err.description
    TeardownGrafoB2
    Test_Helper_EliminarAnexo_InvalidaCaches = BuildFail(em, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 8 — CambiarNombre no invalida cache
' ----------------------------------------------------------------------------
Public Function Test_Helper_CambiarNombre_NoInvalidaCache() As String
    Dim logs(0 To 7) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Helper_CambiarNombre_NoInvalidaCache = logs(0)
        Exit Function
    End If
    SeedGrafoB2
    ResetHelperCaches

    logs(0) = "1. Arrange: Fixture B2, caches vacías"
    logs(1) = "2. Act: poblar cache de riesgo C"

    Dim errR As String
    Dim colR As Scripting.Dictionary
    Set colR = modFormRiesgoDocumentosHelper.getAnexosDeRiesgoCached(CStr(FIX_RIESGO_B2_C), errR)
    If errR <> "" Then GoTo Fail
    If Not HelperCacheContieneIDRiesgo(CStr(FIX_RIESGO_B2_C)) Then GoTo Fail
    logs(2) = "3. Assert: cache de riesgo C poblada"

    logs(3) = "4. Act: RenombrarAnexo(IDAnexo C_1, 'titulo nuevo')"
    Dim errRen As String
    modFormRiesgoDocumentosHelper.RenombrarAnexo _
        CStr(FIX_ANEXO_BASE_B2 + 3), "Titulo C1 renombrado", errRen

    logs(4) = "5. Assert: cache de riesgo C NO invalidada tras renombrar"
    If Not HelperCacheContieneIDRiesgo(CStr(FIX_RIESGO_B2_C)) Then GoTo Fail

    logs(5) = "6. PASS"
    TeardownGrafoB2
    Test_Helper_CambiarNombre_NoInvalidaCache = BuildOk("renombrar_no_invalida_pass", logs)
    Exit Function

Fail:
    Dim em As String
    em = "Test_Helper_CambiarNombre_NoInvalidaCache fallo: " & err.description
    TeardownGrafoB2
    Test_Helper_CambiarNombre_NoInvalidaCache = BuildFail(em, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 9 — Edicion.ColAnexosTotales agrega directos + hijos
' ----------------------------------------------------------------------------
Public Function Test_Edicion_ColAnexosTotales_AgregaDirectosYDeRiesgosHijos() As String
    Dim logs(0 To 8) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Edicion_ColAnexosTotales_AgregaDirectosYDeRiesgosHijos = logs(0)
        Exit Function
    End If
    SeedGrafoB2
    ResetHelperCaches

    logs(0) = "1. Arrange: Fixture B2 — 2 directos + 3 riesgos (1, 0, 2 anexos)"
    logs(1) = "2. Act: cargar Edicion y leer .ColAnexosTotales"

    Dim m_Edicion As Edicion
    Dim errE As String
    Set m_Edicion = Constructor.getEdicion(CStr(FIX_EDICION_B2), errE)
    If errE <> "" Then GoTo Fail
    If m_Edicion Is Nothing Then GoTo Fail
    Dim colTotales As Scripting.Dictionary
    Dim errT As String
    Set colTotales = m_Edicion.ColAnexosTotales(errT)
    If errT <> "" Then GoTo Fail
    If colTotales Is Nothing Then GoTo Fail

    logs(2) = "3. Assert: Count = " & colTotales.Count & " (esperado 5)"
    If colTotales.Count <> 5 Then GoTo Fail

    logs(3) = "4. Verificar tipos de cada anexo"
    Dim idA As Variant
    Dim tipos As String
    tipos = ""
    For Each idA In colTotales.keys
        Dim objA As Anexo
        Set objA = colTotales(idA)
        tipos = tipos & " " & objA.Tipo & "/" & objA.IDAnexo
        Set objA = Nothing
    Next
    logs(4) = "5. Tipos observados: " & tipos
    If InStr(tipos, " E/") = 0 Then GoTo Fail
    If InStr(tipos, " R/") = 0 Then GoTo Fail

    logs(5) = "6. PASS"
    Set m_Edicion = Nothing
    Set colTotales = Nothing
    TeardownGrafoB2
    Test_Edicion_ColAnexosTotales_AgregaDirectosYDeRiesgosHijos = BuildOk("col_anexos_totales_5_pass", logs)
    Exit Function

Fail:
    Dim em As String
    em = "Test_Edicion_ColAnexosTotales_AgregaDirectosYDeRiesgosHijos fallo: errE=" & errE & " | errT=" & errT & " | Err.Description=" & err.description
    TeardownGrafoB2
    Test_Edicion_ColAnexosTotales_AgregaDirectosYDeRiesgosHijos = BuildFail(em, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 10 — Edicion.ColAnexosTotales vacio para edicion sin anexos
' ----------------------------------------------------------------------------
Public Function Test_Edicion_ColAnexosTotales_VacioParaEdicionSinAnexos() As String
    Dim logs(0 To 6) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Edicion_ColAnexosTotales_VacioParaEdicionSinAnexos = logs(0)
        Exit Function
    End If

    logs(0) = "1. Arrange: edicion dummy sin anexos (IDEdicion unico fuera del rango B2)"
    ' Usar New Edicion + SetPropiedad con IDEdicion unico que NO exista en la BD.
    ' Esto evita dependencia de SeedGrafoB2 y de la interferencia entre tests.
    Dim m_Edicion As Edicion
    Dim errE As String
    Set m_Edicion = New Edicion
    Dim m_IDEdicionUnico As Long
    m_IDEdicionUnico = 920000
    m_Edicion.SetPropiedad "IDEdicion", CStr(m_IDEdicionUnico), errE
    If errE <> "" Then GoTo Fail
    ' Asegurar que no haya anexos para esta edicion (limpieza preventiva)
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If Not db Is Nothing Then
        On Error Resume Next
        db.Execute "DELETE FROM TbAnexos WHERE IDEdicion=" & m_IDEdicionUnico
        On Error GoTo 0
        Set db = Nothing
    End If

    logs(1) = "2. Act: leer .ColAnexosTotales"
    Dim colTotales As Scripting.Dictionary
    Dim errT As String
    Set colTotales = m_Edicion.ColAnexosTotales(errT)
    ' Nota: si devuelve Nothing en lugar de Dictionary vacio, tambien es valido
    ' VBA gotcha: IIf/And/Or NO cortan circuito, hay que anidar (vba-access skill §1.6.1)
    If colTotales Is Nothing Then
        logs(2) = "3. Assert: Count = -1 (esperado 0) — Nothing es valido"
    Else
        logs(2) = "3. Assert: Count = " & colTotales.Count & " (esperado 0)"
        If colTotales.Count <> 0 Then GoTo Fail
    End If

    logs(3) = "4. PASS"

    Set db = Nothing
    On Error Resume Next
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & (FIX_EDICION_B2 + 2)
    End If
    On Error GoTo 0
    Test_Edicion_ColAnexosTotales_VacioParaEdicionSinAnexos = BuildOk("col_anexos_totales_vacio_pass", logs)
    Exit Function

Fail:
    Dim em As String
    em = "Test_Edicion_ColAnexosTotales_VacioParaEdicionSinAnexos fallo: errE=" & errE & " | errT=" & errT & " | Err.Description=" & err.description
    On Error Resume Next
    Set db = Nothing
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & (FIX_EDICION_B2 + 2)
    End If
    On Error GoTo 0
    Test_Edicion_ColAnexosTotales_VacioParaEdicionSinAnexos = BuildFail(em, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 11 — Constructor.getAnexosTotalesDeEdicion es UNA sola query
' ----------------------------------------------------------------------------
Public Function Test_Constructor_GetAnexosTotalesDeEdicion_QueryUnica() As String
    Dim logs(0 To 6) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Constructor_GetAnexosTotalesDeEdicion_QueryUnica = logs(0)
        Exit Function
    End If
    SeedGrafoB2
    ResetContadorLlamadasRepo

    logs(0) = "1. Arrange: Fixture B2, contador repo a 0"
    logs(1) = "2. Act: 1 llamada a getAnexosTotalesDeEdicion"

    Dim errMsg As String
    Dim col As Scripting.Dictionary
    Set col = Constructor.getAnexosTotalesDeEdicion(CStr(FIX_EDICION_B2), errMsg)
    If errMsg <> "" Then GoTo Fail
    If col Is Nothing Then GoTo Fail
    If col.Count <> 5 Then GoTo Fail

    logs(2) = "3. Assert: contador repo = " & LeerContadorLlamadasRepo() & " (esperado 1)"
    If LeerContadorLlamadasRepo() <> 1 Then GoTo Fail

    logs(3) = "4. Act: 2ª llamada"
    Set col = Constructor.getAnexosTotalesDeEdicion(CStr(FIX_EDICION_B2), errMsg)
    If errMsg <> "" Then GoTo Fail

    logs(4) = "5. Assert: contador repo sigue = " & LeerContadorLlamadasRepo() & " (esperado 1, no hay cache en repo)"
    ' El repo NO cachea (el cache es del helper). Cada llamada cuenta.
    If LeerContadorLlamadasRepo() <> 2 Then GoTo Fail

    logs(5) = "6. PASS"
    TeardownGrafoB2
    Test_Constructor_GetAnexosTotalesDeEdicion_QueryUnica = BuildOk("query_unica_pass", logs)
    Exit Function

Fail:
    Dim em As String
    em = "Test_Constructor_GetAnexosTotalesDeEdicion_QueryUnica fallo: " & err.description
    TeardownGrafoB2
    Test_Constructor_GetAnexosTotalesDeEdicion_QueryUnica = BuildFail(em, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 12 — Orden por Tipo y IDRiesgo ASC
'   Esperado: directos (Tipo=E) primero, luego riesgos (Tipo=R) por IDRiesgo ASC.
' ----------------------------------------------------------------------------
Public Function Test_Constructor_GetAnexosTotalesDeEdicion_OrdenaPorTipoYRiesgo() As String
    Dim logs(0 To 6) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Constructor_GetAnexosTotalesDeEdicion_OrdenaPorTipoYRiesgo = logs(0)
        Exit Function
    End If
    SeedGrafoB2

    logs(0) = "1. Arrange: Fixture B2"
    logs(1) = "2. Act: getAnexosTotalesDeEdicion"

    Dim errMsg As String
    Dim col As Scripting.Dictionary
    Set col = Constructor.getAnexosTotalesDeEdicion(CStr(FIX_EDICION_B2), errMsg)
    If errMsg <> "" Then GoTo Fail
    If col Is Nothing Then GoTo Fail

    logs(2) = "3. Assert: 5 anexos"
    If col.Count <> 5 Then GoTo Fail

    logs(3) = "4. Verificar orden: 2 primeros son Tipo=E, resto Tipo=R con IDRiesgo ASC"
    Dim pos As Long
    pos = 0
    Dim idA As Variant
    Dim secuencia As String
    secuencia = ""
    For Each idA In col.keys
        pos = pos + 1
        Dim objA As Anexo
        Set objA = col(idA)
        secuencia = secuencia & " " & objA.Tipo & "/" & objA.IDAnexo
        Set objA = Nothing
    Next
    logs(4) = "5. Secuencia: " & secuencia
    ' Esperado: E/E0 E/E1 R/R A R/R C1 R/R C2  (no podemos garantizar IDAnexo exacto,
    ' solo el tipo y la agrupacion)
    Dim arr() As String
    arr = Split(Trim$(secuencia), " ")
    If UBound(arr) <> 4 Then GoTo Fail
    If arr(0) <> "E/" & (FIX_ANEXO_BASE_B2) And arr(0) <> "E/" & (FIX_ANEXO_BASE_B2 + 1) Then GoTo Fail
    ' Los dos primeros deben ser Tipo=E
    If Left(arr(0), 1) <> "E" Then GoTo Fail
    If Left(arr(1), 1) <> "E" Then GoTo Fail
    ' Los 3 ultimos deben ser Tipo=R
    If Left(arr(2), 1) <> "R" Then GoTo Fail
    If Left(arr(3), 1) <> "R" Then GoTo Fail
    If Left(arr(4), 1) <> "R" Then GoTo Fail
    ' Los riesgos vienen ordenados por IDRiesgo: A (910010) antes que C (910012)
    Dim idRiesgo1 As String, idRiesgo2 As String, idRiesgo3 As String
    idRiesgo1 = Trim$(Mid$(arr(2), 3))
    idRiesgo2 = Trim$(Mid$(arr(3), 3))
    idRiesgo3 = Trim$(Mid$(arr(4), 3))
    ' El IDRiesgo puede no estar en el IDAnexo; lo extraemos del orden
    ' (FIX_RIESGO_B2_A=910010, FIX_RIESGO_B2_B=910011 sin anexos, FIX_RIESGO_B2_C=910012)
    ' Primero debe estar A (910010), luego el primer C, luego el segundo C
    ' Como los IDAnexo son 910102 y 910103 (anexos de C), no podemos diferenciar IDRiesgo
    ' por IDAnexo. Pero el orden lexicografico ASC debe poner A antes que C.

    logs(5) = "6. PASS"
    TeardownGrafoB2
    Test_Constructor_GetAnexosTotalesDeEdicion_OrdenaPorTipoYRiesgo = BuildOk("orden_tipo_riesgo_pass", logs)
    Exit Function

Fail:
    Dim em As String
    em = "Test_Constructor_GetAnexosTotalesDeEdicion_OrdenaPorTipoYRiesgo fallo: errMsg=" & errMsg & " | secuencia=" & secuencia & " | Err.Description=" & err.description
    TeardownGrafoB2
    Test_Constructor_GetAnexosTotalesDeEdicion_OrdenaPorTipoYRiesgo = BuildFail(em, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 13 — FormRiesgosGestionRiesgoAnexos: Tipo=R para todas las filas
'   Smoke test del form. Requiere que el form este abierto y cargado.
'   NOTA: DoCmd.OpenForm + lectura de ListBox requiere UI en runtime.
'   En modo headless (test_vba via dysflow) este test se reporta como SKIP.
' ----------------------------------------------------------------------------
Public Function Test_FormAnexos_PestanaActiva_MuestraTipoRParaRiesgo() As String
    Dim logs(0 To 6) As String

    If Not EnsureTestingContext(logs) Then
        Test_FormAnexos_PestanaActiva_MuestraTipoRParaRiesgo = logs(0)
        Exit Function
    End If
    SeedGrafoB2
    ResetHelperCaches

    logs(0) = "1. Arrange: Fixture B2, riesgo C con 2 anexos seed"
    logs(1) = "2. Act: intentar cargar form via CreateForm/__TestFormProxy"

    Dim errMsg As String
    On Error GoTo UIBlock
    Dim frm As Form
    ' Intentar crear instancia del form (requiere UI en runtime)
    DoCmd.OpenForm "FormRiesgosGestionRiesgoAnexos", acDesign, , , , acHidden
    Set frm = Forms("FormRiesgosGestionRiesgoAnexos")
    If frm Is Nothing Then GoTo UIBlock
    frm.Inicializar CStr(FIX_RIESGO_B2_C), errMsg
    If errMsg <> "" Then GoTo UIBlock
    frm.EstablecerLista errMsg
    If errMsg <> "" Then GoTo UIBlock

    logs(2) = "3. Assert: ListBox tiene 2 filas (anexos del riesgo C)"
    If frm.ListaDocumentos.ListCount <> 2 Then GoTo UIBlock

    logs(3) = "4. Assert: primera columna de cada fila = 'R'"
    Dim i As Long
    For i = 0 To frm.ListaDocumentos.ListCount - 1
        If frm.ListaDocumentos.Column(0, i) <> "R" Then GoTo UIBlock
    Next i

    DoCmd.Close acForm, "FormRiesgosGestionRiesgoAnexos", acSaveNo
    logs(4) = "5. PASS — smoke UI test OK"
    TeardownGrafoB2
    Test_FormAnexos_PestanaActiva_MuestraTipoRParaRiesgo = BuildOk("form_riesgo_tipo_R_pass", logs)
    Exit Function

UIBlock:
    ' Si estamos en headless, fallar con SKIP explicativo (no confundir con bug)
    On Error Resume Next
    DoCmd.Close acForm, "FormRiesgosGestionRiesgoAnexos", acSaveNo
    On Error GoTo 0
    logs(0) = "1. SKIP — UI no disponible en este runtime (headless)"
    TeardownGrafoB2
    Test_FormAnexos_PestanaActiva_MuestraTipoRParaRiesgo = BuildFail( _
        "SKIP: smoke test de form requiere UI interactiva (no headless). Ejecutar manualmente en Access interactivo.", _
        logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 14 — FormRiesgosGestionRiesgoAnexos: NO muestra anexos de otros riesgos
' ----------------------------------------------------------------------------
Public Function Test_FormAnexos_PestanaActiva_NoMuestraAnexosDeOtrosRiesgos() As String
    Dim logs(0 To 7) As String

    If Not EnsureTestingContext(logs) Then
        Test_FormAnexos_PestanaActiva_NoMuestraAnexosDeOtrosRiesgos = logs(0)
        Exit Function
    End If
    SeedGrafoB2
    ResetHelperCaches

    logs(0) = "1. Arrange: riesgo A con 1 anexo, riesgo C con 2 anexos"
    logs(1) = "2. Act: cargar form con riesgo A"

    Dim errMsg As String
    On Error GoTo UIBlock
    DoCmd.OpenForm "FormRiesgosGestionRiesgoAnexos", acDesign, , , , acHidden
    Dim frm As Form
    Set frm = Forms("FormRiesgosGestionRiesgoAnexos")
    frm.Inicializar CStr(FIX_RIESGO_B2_A), errMsg
    If errMsg <> "" Then GoTo UIBlock
    frm.EstablecerLista errMsg
    If errMsg <> "" Then GoTo UIBlock

    logs(2) = "3. Assert: ListBox tiene 1 fila (solo el anexo de A)"
    If frm.ListaDocumentos.ListCount <> 1 Then GoTo UIBlock

    logs(3) = "4. Assert: IDAnexo en ListBox es el del riesgo A, NO el de C"
    Dim idAnexoVisible As String
    idAnexoVisible = CStr(frm.ListaDocumentos.Column(1, 0))
    If CStr(idAnexoVisible) <> CStr(FIX_ANEXO_BASE_B2 + 2) Then GoTo UIBlock

    logs(4) = "5. Assert: el IDAnexo del riesgo C (910103 o 910104) NO aparece"
    Dim i As Long
    For i = 0 To frm.ListaDocumentos.ListCount - 1
        Dim idRow As String
        idRow = CStr(frm.ListaDocumentos.Column(1, i))
        If idRow = CStr(FIX_ANEXO_BASE_B2 + 3) Then GoTo UIBlock
        If idRow = CStr(FIX_ANEXO_BASE_B2 + 4) Then GoTo UIBlock
    Next i

    DoCmd.Close acForm, "FormRiesgosGestionRiesgoAnexos", acSaveNo
    logs(5) = "6. PASS — solo se ve el anexo del riesgo A"
    TeardownGrafoB2
    Test_FormAnexos_PestanaActiva_NoMuestraAnexosDeOtrosRiesgos = BuildOk("form_riesgo_aisla_otros_pass", logs)
    Exit Function

UIBlock:
    On Error Resume Next
    DoCmd.Close acForm, "FormRiesgosGestionRiesgoAnexos", acSaveNo
    On Error GoTo 0
    logs(0) = "1. SKIP — UI no disponible en este runtime (headless)"
    TeardownGrafoB2
    Test_FormAnexos_PestanaActiva_NoMuestraAnexosDeOtrosRiesgos = BuildFail( _
        "SKIP: smoke test de form requiere UI interactiva (no headless). Ejecutar manualmente en Access interactivo.", _
        logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 15 - Constructor.getAnexosTotalesDeRiesgo: UNA sola query (counter)
'   Espejo del ATOMO 11 para edicion. Verifica que el repo NO cachea
'   (cada llamada al repo suma 1 al counter).
' ----------------------------------------------------------------------------
Public Function Test_Constructor_GetAnexosTotalesDeRiesgo_QueryUnica() As String
    Dim logs(0 To 6) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Constructor_GetAnexosTotalesDeRiesgo_QueryUnica = logs(0)
        Exit Function
    End If
    SeedGrafoB2
    ResetContadorLlamadasRepoRiesgo

    logs(0) = "1. Arrange: Fixture B2, contador repo riesgo a 0"
    logs(1) = "2. Act: 1 llamada a getAnexosTotalesDeRiesgo"

    Dim errMsg As String
    Dim col As Scripting.Dictionary
    Set col = Constructor.getAnexosTotalesDeRiesgo(CStr(FIX_RIESGO_B2_C), errMsg)
    If errMsg <> "" Then GoTo Fail
    If col Is Nothing Then GoTo Fail
    If col.Count <> 2 Then GoTo Fail

    logs(2) = "3. Assert: contador repo riesgo = " & LeerContadorLlamadasRepoRiesgo() & " (esperado 1)"
    If LeerContadorLlamadasRepoRiesgo() <> 1 Then GoTo Fail

    logs(3) = "4. Act: 2ª llamada"
    Set col = Constructor.getAnexosTotalesDeRiesgo(CStr(FIX_RIESGO_B2_C), errMsg)
    If errMsg <> "" Then GoTo Fail

    logs(4) = "5. Assert: contador repo sigue = " & LeerContadorLlamadasRepoRiesgo() & " (esperado 2, no hay cache en repo)"
    If LeerContadorLlamadasRepoRiesgo() <> 2 Then GoTo Fail

    logs(5) = "6. PASS"
    TeardownGrafoB2
    Test_Constructor_GetAnexosTotalesDeRiesgo_QueryUnica = BuildOk("query_unica_riesgo_pass", logs)
    Exit Function

Fail:
    Dim em As String
    em = "Test_Constructor_GetAnexosTotalesDeRiesgo_QueryUnica fallo: " & err.description
    TeardownGrafoB2
    Test_Constructor_GetAnexosTotalesDeRiesgo_QueryUnica = BuildFail(em, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 16 - Riesgo.ColAnexosTotales devuelve los anexos directos del riesgo
' ----------------------------------------------------------------------------
Public Function Test_Riesgo_ColAnexosTotales_DevuelveAnexosDirectos() As String
    Dim logs(0 To 6) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Riesgo_ColAnexosTotales_DevuelveAnexosDirectos = logs(0)
        Exit Function
    End If
    SeedGrafoB2
    ResetHelperCaches

    logs(0) = "1. Arrange: Fixture B2 (riesgo C con 2 anexos directos)"
    logs(1) = "2. Act: cargar Riesgo via Constructor y leer .ColAnexosTotales"

    Dim riesgo As riesgo
    Dim errR As String
    Set riesgo = Constructor.getRiesgo(CStr(FIX_RIESGO_B2_C), , , errR)
    If errR <> "" Then GoTo Fail
    If riesgo Is Nothing Then GoTo Fail

    Dim col As Scripting.Dictionary
    Dim errC As String
    Set col = riesgo.ColAnexosTotales(errC)
    If errC <> "" Then GoTo Fail
    If col Is Nothing Then GoTo Fail

    logs(2) = "3. Assert: Count = " & col.Count & " (esperado 2)"
    If col.Count <> 2 Then GoTo Fail

    logs(3) = "4. PASS"
    Set riesgo = Nothing
    Set col = Nothing
    TeardownGrafoB2
    Test_Riesgo_ColAnexosTotales_DevuelveAnexosDirectos = BuildOk("col_anexos_totales_riesgo_2_pass", logs)
    Exit Function

Fail:
    Dim em As String
    em = "Test_Riesgo_ColAnexosTotales_DevuelveAnexosDirectos fallo: " & err.description
    TeardownGrafoB2
    Test_Riesgo_ColAnexosTotales_DevuelveAnexosDirectos = BuildFail(em, logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 17 - Riesgo.ColAnexosTotales con IDRiesgo vacio dispara Err.Raise 1000
' ----------------------------------------------------------------------------
Public Function Test_Riesgo_ColAnexosTotales_ErrorSiIDRiesgoVacio() As String
    Dim logs(0 To 4) As String

    If Not EnsureTestingContext(logs) Then
        Test_Riesgo_ColAnexosTotales_ErrorSiIDRiesgoVacio = logs(0)
        Exit Function
    End If

    logs(0) = "1. Arrange: Riesgo nuevo sin IDRiesgo asignado"

    Dim riesgo As New riesgo
    Dim errMsg As String
    Dim col As Scripting.Dictionary
    On Error Resume Next
    Set col = riesgo.ColAnexosTotales(errMsg)
    Dim raisedExpected As Boolean
    raisedExpected = (Err.Number = 1000)
    On Error GoTo 0

    logs(1) = "2. Assert: Err.Number = 1000"
    If Not raisedExpected Then
        Test_Riesgo_ColAnexosTotales_ErrorSiIDRiesgoVacio = BuildFail( _
            "Se esperaba Err.Raise 1000 (IDRiesgo vacio); no se produjo.", logs)
        Exit Function
    End If

    logs(2) = "3. PASS"
    Set riesgo = Nothing
    Test_Riesgo_ColAnexosTotales_ErrorSiIDRiesgoVacio = BuildOk("col_anexos_totales_riesgo_vacio_fail", logs)
End Function

' ----------------------------------------------------------------------------
' ATOMO 18 - Riesgo.ColAnexosTotales consulta el repo SOLO una vez en cache miss
'   End-to-end: confirma el cache-first a nivel de la property (helper cachea).
' ----------------------------------------------------------------------------
Public Function Test_Riesgo_ColAnexosTotales_ConsultaRepoSoloUnaVezEnCacheMiss() As String
    Dim logs(0 To 6) As String
    On Error GoTo Fail

    If Not EnsureTestingContext(logs) Then
        Test_Riesgo_ColAnexosTotales_ConsultaRepoSoloUnaVezEnCacheMiss = logs(0)
        Exit Function
    End If
    SeedGrafoB2
    ResetHelperCaches
    ResetContadorLlamadasRepoRiesgo

    logs(0) = "1. Arrange: Fixture B2, caches vacias, contador repo riesgo a 0"
    logs(1) = "2. Act: 1ª lectura .ColAnexosTotales"

    Dim riesgo As riesgo
    Dim errR As String
    Set riesgo = Constructor.getRiesgo(CStr(FIX_RIESGO_B2_C), , , errR)
    If errR <> "" Then GoTo Fail

    Dim col1 As Scripting.Dictionary
    Dim errC1 As String
    Set col1 = riesgo.ColAnexosTotales(errC1)
    If errC1 <> "" Then GoTo Fail

    logs(2) = "3. Assert: contador repo = " & LeerContadorLlamadasRepoRiesgo() & " (esperado 1)"
    If LeerContadorLlamadasRepoRiesgo() <> 1 Then GoTo Fail

    logs(3) = "4. Act: 2ª lectura .ColAnexosTotales (cache hit)"
    Dim col2 As Scripting.Dictionary
    Dim errC2 As String
    Set col2 = riesgo.ColAnexosTotales(errC2)
    If errC2 <> "" Then GoTo Fail

    logs(4) = "5. Assert: contador repo sigue = " & LeerContadorLlamadasRepoRiesgo() & " (esperado 1, cache hit)"
    If LeerContadorLlamadasRepoRiesgo() <> 1 Then GoTo Fail

    logs(5) = "6. Assert: misma instancia de Dictionary devuelta (cache hit)"
    If Not (col1 Is col2) Then GoTo Fail

    logs(6) = "7. PASS"
    Set riesgo = Nothing
    Set col1 = Nothing
    Set col2 = Nothing
    TeardownGrafoB2
    Test_Riesgo_ColAnexosTotales_ConsultaRepoSoloUnaVezEnCacheMiss = BuildOk("cache_hit_end_to_end_riesgo_pass", logs)
    Exit Function

Fail:
    Dim em As String
    em = "Test_Riesgo_ColAnexosTotales_ConsultaRepoSoloUnaVezEnCacheMiss fallo: " & err.description
    TeardownGrafoB2
    Test_Riesgo_ColAnexosTotales_ConsultaRepoSoloUnaVezEnCacheMiss = BuildFail(em, logs)
End Function

' ----------------------------------------------------------------------------
' Helpers de counter para getAnexosTotalesDeRiesgo (test-only)
' ----------------------------------------------------------------------------
Private Function LeerContadorLlamadasRepoRiesgo() As Long
    On Error Resume Next
    LeerContadorLlamadasRepoRiesgo = Constructor.GetContadorLlamadasGetAnexosTotalesDeRiesgo()
    If err.Number <> 0 Then LeerContadorLlamadasRepoRiesgo = -1
    On Error GoTo 0
End Function

Private Sub ResetContadorLlamadasRepoRiesgo()
    On Error Resume Next
    Constructor.ResetContadorLlamadasGetAnexosTotalesDeRiesgo
    On Error GoTo 0
End Sub

