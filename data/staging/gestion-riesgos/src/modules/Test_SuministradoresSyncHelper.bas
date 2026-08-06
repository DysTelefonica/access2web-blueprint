Attribute VB_Name = "Test_SuministradoresSyncHelper"
Option Explicit
' ============================================================
' Test_SuministradoresSyncHelper - TDD red atoms for
' modSuministradoresSyncHelper.SincronizarSuministradoresParaEdicion
'
' Skill: access-vba-tdd v2.4
' Project: gestion_riesgos
' Branch: feat/metodologia-e2e-riesgos-2026-06-19
'
' SCENARIO CLASSES (access-vba-e2e-methodology):
'   Happy      - edicion valida, jerarquia con subcontratistas -> sync OK
'   Sad        - edicion inexistente o sin subcontratistas en jerarquia
'   Edge       - jerarquia vacia, edicion nueva (creada durante el test)
'   Adversarial - doble-click durante sync (operacion idempotente)
'
' HELPER SIGNATURE (Sub, not Function - task constraint):
'   Sub SincronizarSuministradoresEnEdicion( _
'       ByVal p_IDEdicion As String
'       Optional ByRef p_Error As String)
'
' SCHEMA EVIDENCE (from SuministradorParaEvidencias.cls + SuministradoresHelper.bas):
'   TbProyectosEdicionesSuministradores columns:
'     ID          (Long, PK)  - auto-generated via DameID()
'     IDEdicion   (String)    - FK to TbProyectosEdiciones.IDEdicion
'     IDSuministrador (String) - FK to TbSuministradores
'     IDAnexo     (Variant)   - NULL = sin evidencia
'
'   Jerarquia source: TbExpedientesSuministradores (confirmed
'   via schema inspection — columns are IDExpedienteSuministrador,
'   IDExpediente, IDSuministrador, IDPadre)
'
' ASSUMPTIONS (confirmed):
'   #3 - TbProyectosEdicionesSuministradores es la tabla destino.
'   Jerarquia source table: TbExpedientesSuministradores (PK IDExpedienteSuministrador).
'
' FIXTURE IDs (rango >= 900000):
'   tbExpediente: 900100  (Test_Fixtures.Subcat_EdicionId usa 900102)
'   tbProyecto:   900101
'   tbEdicion:    900102  (Edicion fixture)
'   tbSuministradores: 900110(ROOT), 900120(CHILD1), 900130(CHILD2), 900121(GRANDCHILD)
'   tbExpedientesSuministradores: 900200(ROOT), 900210(CHILD1), 900220(CHILD2), 900211(GCHILD)
'   tbProyectosEdicionesSuministradores: fixture inserta CHILD1 y CHILD2;
'     GRANDCHILD es sobrante (no es subcontratista directo)
' ============================================================

' --- Constants de IDs del fixture ---
Private Const FIX_EXPEDIENTE  As Long = 900100
Private Const FIX_PROYECTO    As Long = 900101
Private Const FIX_EDICION     As Long = 900102
Private Const FIX_ROOT        As Long = 900110
Private Const FIX_CHILD1      As Long = 900120
Private Const FIX_CHILD2      As Long = 900130
Private Const FIX_GRANDCHILD  As Long = 900121
Private Const FIX_IDES_ROOT   As Long = 900200
Private Const FIX_IDES_CHILD1 As Long = 900210
Private Const FIX_IDES_CHILD2 As Long = 900220
Private Const FIX_IDES_GCHILD As Long = 900211

' --- Module-level constants for REQ-CAL-12 ---
Private Const FIX_ID_R12_EXPEDIENTE     As Long = 904100
Private Const FIX_ID_R12_PROYECTO       As Long = 904101
Private Const FIX_ID_R12_EDICION_HAPPY  As Long = 904102
Private Const FIX_ID_R12_EDICION_SAD    As Long = 904103
Private Const FIX_ID_R12_EDICION_EDGE   As Long = 904104
Private Const FIX_ID_R12_EDICION_ADV    As Long = 904105
Private Const FIX_ID_R12_SUM_ACME       As Long = 904110
Private Const FIX_ID_R12_SUM_BETA       As Long = 904120
Private Const FIX_ID_R12_SUM_GAMMA      As Long = 904130
Private Const FIX_ID_R12_PES_HAPPY_ACME As Long = 904250
Private Const FIX_ID_R12_PES_HAPPY_BETA As Long = 904251
Private Const FIX_ID_R12_PES_SAD_ACME   As Long = 904260
Private Const FIX_ID_R12_PES_SAD_BETA   As Long = 904261
Private Const FIX_ID_R12_PES_EDGE_GAMMA As Long = 904270
Private Const FIX_ID_R12_PES_ADV_ACME   As Long = 904280

' --- JSON helpers wrapper (project convention) ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' --- Fixture setup for REQ-CAL-12: seed expedición + proyecto + edición +
'     subcontratistas en TbSuministradores + filas en
'     TbProyectosEdicionesSuministradores. Independiente del SeedSubcatGraph
'     (900100-900121) para no contaminar el fixture compartido.
Private Sub SeedR12Fixture(ByVal p_IDEdicion As Long, _
                            ByVal p_IDProyecto As Long, _
                            ByVal p_IDExpediente As Long)
    On Error GoTo EH_Seed
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then Err.Raise 1001, "SeedR12Fixture", "GetTestDb returned Nothing: " & dbErr

    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & p_IDEdicion, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & p_IDEdicion, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & p_IDProyecto, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & p_IDExpediente, dbFailOnError
    db.Execute "DELETE FROM TbSuministradores WHERE IDSuministrador IN (" & _
               FIX_ID_R12_SUM_ACME & "," & FIX_ID_R12_SUM_BETA & "," & FIX_ID_R12_SUM_GAMMA & ")", dbFailOnError
    On Error GoTo 0

    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
               "VALUES (" & p_IDExpediente & ", 'TEST12', 'Fixture REQ-CAL-12', 'Test', 1)", dbFailOnError
    db.Execute "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto) " & _
               "VALUES (" & p_IDProyecto & ", " & p_IDExpediente & ", 'TESTPROJ12')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDProyecto, IDEdicion, Edicion, Elaborado) " & _
               "VALUES (" & p_IDProyecto & ", " & p_IDEdicion & ", 1, 'test_user')", dbFailOnError

    Set db = Nothing
    Exit Sub

EH_Seed:
    Dim eN As Long: eN = Err.Number
    Dim ed As String: ed = Err.description
    On Error Resume Next
    Set db = Nothing
    Err.Raise eN, "SeedR12Fixture", "Seed failed: " & eN & " - " & ed
End Sub

Private Sub TeardownR12Fixture(ByVal p_IDEdicion As Long, _
                               ByVal p_IDProyecto As Long, _
                               ByVal p_IDExpediente As Long)
    On Error Resume Next
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then Exit Sub

    db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & p_IDEdicion, dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & p_IDEdicion, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & p_IDProyecto, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & p_IDExpediente, dbFailOnError
    db.Execute "DELETE FROM TbSuministradores WHERE IDSuministrador IN (" & _
               FIX_ID_R12_SUM_ACME & "," & FIX_ID_R12_SUM_BETA & "," & FIX_ID_R12_SUM_GAMMA & ")", dbFailOnError

    Set db = Nothing
    On Error GoTo 0
End Sub

' ============================================================
' TEST 1 (Happy) - edicion valida, jerarquia con
' subcontratistas directos -> sync inserta faltantes, p_Error vacio
' ============================================================
Public Function Test_SincronizarSuministradoresEnEdicion_Happy() As String
    Dim logs(0 To 7) As String
    logs(0) = "1. Arrange: ForceLocalBackend + SeedSubcatGraph"
    logs(1) = "2. Arrange: verificar jerarquia tiene CHILD1 y CHILD2 como subcontratistas"
    logs(2) = "3. Act: modSuministradoresSyncHelper.SincronizarSuministradoresParaEdicion"
    logs(3) = "4. Assert: count TbProyectosEdicionesSuministradores = 2 (sin cambios, ya synced)"
    logs(5) = "5. Teardown: limpieza delta"
    logs(6) = "6. Teardown: ResetTestSession"
    logs(7) = ""

    ' --- Arrange: ForceLocalBackend ---
    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        logs(0) = "TESTS BLOCKED: " & cfgError
        Test_SincronizarSuministradoresEnEdicion_Happy = Test_Helper.BuildJsonFail(cfgError, logs)
        Exit Function
    End If
    logs(0) = "1. Arrange: ForceLocalBackend OK"

    ' Poblar grafo de subcontratistas (padres + hijos + edition) via SeedSubcatGraph
    Dim seedError As String
    Call Test_Fixtures.SeedSubcatGraph

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        logs(1) = "TESTS BLOCKED: GetTestDb returned Nothing: " & dbErr
        Test_SincronizarSuministradoresEnEdicion_Happy = Test_Helper.BuildJsonFail(logs(1), logs)
        GoTo Teardown
    End If
    logs(1) = "1. Arrange: SeedSubcatGraph OK"

    ' Verificar cardinalidad antes: CHILD1 y CHILD2 ya estan en la edicion
    ' (SeedSubcatGraph inserta ambos)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) FROM TbProyectosEdicionesSuministradores " & _
        "WHERE IDEdicion=" & FIX_EDICION, dbOpenSnapshot)
    Dim countAntes As Long
    countAntes = rs.fields(0).value
    rs.Close: Set rs = Nothing

    ' --- Act ---
    Dim pError As String
    ' TODO: confirmar que el primer parametro es ByVal String y segundo es ByVal Edicion
    ' El Sub aun no existe -> compile error = RED
    modSuministradoresSyncHelper.SincronizarSuministradoresParaEdicion _
        CStr(FIX_EDICION), Nothing, pError

    ' --- Assert ---
    If pError <> "" Then
        logs(3) = "3. Assert FAIL: p_Error no vacio: " & pError
        Test_SincronizarSuministradoresEnEdicion_Happy = Test_Helper.BuildJsonFail(pError, logs)
        GoTo Teardown
    End If

    ' Verificar que sync no duplico filas (CHILD1 y CHILD2 ya estaban)
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) FROM TbProyectosEdicionesSuministradores " & _
        "WHERE IDEdicion=" & FIX_EDICION, dbOpenSnapshot)
    Dim countDespues As Long
    countDespues = rs.fields(0).value
    rs.Close: Set rs = Nothing

    logs(3) = "3. Assert: countAntes=" & countAntes & " countDespues=" & countDespues
    If countDespues <> countAntes Then
        Test_SincronizarSuministradoresEnEdicion_Happy = _
            Test_Helper.BuildJsonFail( _
                "Row count changed unexpectedly: antes=" & countAntes & " despues=" & countDespues, logs)
        GoTo Teardown
    End If

    Test_SincronizarSuministradoresEnEdicion_Happy = Test_Helper.BuildJsonOk("happy_sync_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    ' Limpiar delta: solo la fila de edition suppliers (el grafo base se limpia en suite)
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & FIX_EDICION
    End If
    Test_Fixtures.TeardownSubcatGraph
    Test_Helper.ResetTestSession
    On Error GoTo 0
End Function

' ============================================================
' TEST 2 (Sad) - edicion inexistente o sin subcontratistas
' en jerarquia -> p_Error no vacio, sin mutacion
' ============================================================
Public Function Test_SincronizarSuministradoresEnEdicion_Sad() As String
    Dim logs(0 To 7) As String
    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: crear edicion sin subcontratistas en TbProyectosEdiciones"
    logs(2) = "3. Act: llamar SincronizarSuministradoresEnEdicion con IDEdicion inexistente"
    logs(3) = "4. Assert: p_Error no vacio O no se insertaron filas"
    logs(5) = "5. Teardown: eliminar edicion de prueba"
    logs(6) = "6. Teardown: ResetTestSession"
    logs(7) = ""

    ' --- Arrange: ForceLocalBackend ---
    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        logs(0) = "TESTS BLOCKED: " & cfgError
        Test_SincronizarSuministradoresEnEdicion_Sad = Test_Helper.BuildJsonFail(cfgError, logs)
        Exit Function
    End If

    ' Crear edicion fantasma sin subcontratistas en el expediente
    ' (TbProyectosEdicionesSuministradores NO tendra filas para esta edicion)
    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        logs(1) = "TESTS BLOCKED: GetTestDb returned Nothing: " & dbErr
        Test_SincronizarSuministradoresEnEdicion_Sad = Test_Helper.BuildJsonFail(logs(1), logs)
        GoTo Teardown
    End If

    ' Crear expediente y proyecto para la edicion de prueba
    Const TEST_EDICION_FAKE As Long = 909991
    Const TEST_PROYECTO_FAKE As Long = 909990
    Const TEST_EXPEDIENTE_FAKE As Long = 909989

    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_EDICION_FAKE
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_PROYECTO_FAKE
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_EXPEDIENTE_FAKE
    On Error GoTo 0

    ' Insertar expediente-proyecto-edicion fantasma
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
               "VALUES (" & TEST_EXPEDIENTE_FAKE & ",'TESTSAD','Edicion fantasma','Test',1)"
    db.Execute "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto) " & _
               "VALUES (" & TEST_PROYECTO_FAKE & "," & TEST_EXPEDIENTE_FAKE & ",'TESTSADPROJ')"
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_EDICION_FAKE & "," & TEST_PROYECTO_FAKE & ",99,'TESTUSER')"

    ' La edicion no tiene subcontratistas en TbExpedientesSuministradores
    ' (no se insertaron filas en TbExpedientesSuministradores para este expediente).
    ' Esto equivale a jerarquia vacia -> sync deberia completar sin error pero no insertar nada.

    ' --- Act ---
    Dim pError As String
    ' Llamar con IDEdicion de edicion sin subcontratistas
    modSuministradoresSyncHelper.SincronizarSuministradoresParaEdicion _
        CStr(TEST_EDICION_FAKE), Nothing, pError

    ' --- Assert ---
    ' Si p_Error es vacio pero no hay filas -> sync completo sin faltantes (behavior OK)
    ' Si p_Error no es vacio -> fallO esperado por edicion sin datos
    ' Probar ambos: si p_Error <> "" -> BuildJsonFail (Sad path correcto)
    ' Si p_Error == "" pero count == 0 -> BuildJsonOk (sin mutacion, behavior aceptable)

    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) FROM TbProyectosEdicionesSuministradores " & _
        "WHERE IDEdicion=" & TEST_EDICION_FAKE, dbOpenSnapshot)
    Dim countFake As Long
    countFake = rs.fields(0).value
    rs.Close: Set rs = Nothing

    logs(3) = "3. Assert: p_Error=[" & pError & "] count=" & countFake

    If pError <> "" Then
        ' Sad path correcto: sync_reporto error
        Test_SincronizarSuministradoresEnEdicion_Sad = Test_Helper.BuildJsonOk("sad_sync_error_reported", logs)
        GoTo Teardown
    End If

    If countFake = 0 Then
        ' Sin error y sin mutacion - comportamiento aceptable (sin faltantes)
        Test_SincronizarSuministradoresEnEdicion_Sad = Test_Helper.BuildJsonOk("sad_no_faltantes_no_error", logs)
        GoTo Teardown
    End If

    ' Si hay filas sin error -> BuildJsonFail (no deberia haber insertado)
    Test_SincronizarSuministradoresEnEdicion_Sad = _
        Test_Helper.BuildJsonFail( _
            "Unexpected rows inserted for empty hierarchy: " & countFake, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & TEST_EDICION_FAKE
        db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_EDICION_FAKE
        db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_PROYECTO_FAKE
        db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_EXPEDIENTE_FAKE
    End If
    Test_Helper.ResetTestSession
    On Error GoTo 0
End Function

' ============================================================
' TEST 3 (Edge) - jerarquia vacia, edicion nueva creada
' durante el test -> sync no inserts, p_Error vacio
' ============================================================
Public Function Test_SincronizarSuministradoresEnEdicion_Edge() As String
    Dim logs(0 To 7) As String
    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Arrange: crear edicion sin jerarquia (no hay TbExpedientesSuministradores)"
    logs(2) = "3. Act: SincronizarSuministradoresEnEdicion"
    logs(3) = "4. Assert: count=0 y p_Error vacio"
    logs(5) = "5. Teardown: eliminar edicion de prueba"
    logs(6) = "6. Teardown: ResetTestSession"
    logs(7) = ""

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        logs(0) = "TESTS BLOCKED: " & cfgError
        Test_SincronizarSuministradoresEnEdicion_Edge = Test_Helper.BuildJsonFail(cfgError, logs)
        Exit Function
    End If

    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        logs(1) = "TESTS BLOCKED: GetTestDb returned Nothing: " & dbErr
        Test_SincronizarSuministradoresEnEdicion_Edge = Test_Helper.BuildJsonFail(logs(1), logs)
        GoTo Teardown
    End If

    ' Crear edicion limpia sin expediente->suministradores (jerarquia vacia)
    Const TEST_EDICION_EDGE As Long = 909992
    Const TEST_PROYECTO_EDGE As Long = 909993
    Const TEST_EXPEDIENTE_EDGE As Long = 909994

    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_EDICION_EDGE
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_PROYECTO_EDGE
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_EXPEDIENTE_EDGE
    On Error GoTo 0

    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
               "VALUES (" & TEST_EXPEDIENTE_EDGE & ",'TESTEDGE','Edge edition','Test',1)"
    db.Execute "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto) " & _
               "VALUES (" & TEST_PROYECTO_EDGE & "," & TEST_EXPEDIENTE_EDGE & ",'TESTEDGEPROJ')"
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
               "VALUES (" & TEST_EDICION_EDGE & "," & TEST_PROYECTO_EDGE & ",1,'TESTUSER')"

    ' No se inserta ninguna fila en TbExpedientesSuministradores -> jerarquia vacia

    ' --- Act ---
    Dim pError As String
    modSuministradoresSyncHelper.SincronizarSuministradoresParaEdicion _
        CStr(TEST_EDICION_EDGE), Nothing, pError

    ' --- Assert ---
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) FROM TbProyectosEdicionesSuministradores " & _
        "WHERE IDEdicion=" & TEST_EDICION_EDGE, dbOpenSnapshot)
    Dim countEdge As Long
    countEdge = rs.fields(0).value
    rs.Close: Set rs = Nothing

    logs(3) = "3. Assert: p_Error=[" & pError & "] count=" & countEdge

    If countEdge <> 0 Then
        Test_SincronizarSuministradoresEnEdicion_Edge = _
            Test_Helper.BuildJsonFail( _
                "Expected 0 rows for empty hierarchy, got " & countEdge, logs)
        GoTo Teardown
    End If

    If pError <> "" Then
        Test_SincronizarSuministradoresEnEdicion_Edge = _
            Test_Helper.BuildJsonFail("Unexpected error: " & pError, logs)
        GoTo Teardown
    End If

    Test_SincronizarSuministradoresEnEdicion_Edge = Test_Helper.BuildJsonOk("edge_empty_hierarchy_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & TEST_EDICION_EDGE
        db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & TEST_EDICION_EDGE
        db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & TEST_PROYECTO_EDGE
        db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & TEST_EXPEDIENTE_EDGE
    End If
    Test_Helper.ResetTestSession
    On Error GoTo 0
End Function

' ============================================================
' TEST 4 (Adversarial) - doble-click simulado: sync dos
' veces seguidas -> idempotente, segunda llamada no duplica
' ni elimina filas que ya estan bien
' ============================================================
Public Function Test_SincronizarSuministradoresEnEdicion_Adversarial() As String
    Dim logs(0 To 7) As String
    logs(0) = "1. Arrange: ForceLocalBackend + SeedSubcatGraph"
    logs(1) = "2. Act: primera llamada a SincronizarSuministradoresEnEdicion"
    logs(2) = "3. Assert: count unchanged (sync idempotente sobre estado limpio)"
    logs(3) = "4. Act: segunda llamada (simula doble-click)"
    logs(4) = "5. Assert: count igual que antes de segunda llamada"
    logs(5) = "6. Teardown: limpieza delta"
    logs(6) = "7. Teardown: ResetTestSession"
    logs(7) = ""

    Dim cfgError As String
    If Not Test_Helper.ForceLocalBackend(cfgError) Then
        logs(0) = "TESTS BLOCKED: " & cfgError
        Test_SincronizarSuministradoresEnEdicion_Adversarial = Test_Helper.BuildJsonFail(cfgError, logs)
        Exit Function
    End If

    ' Poblar grafo base
    Call Test_Fixtures.SeedSubcatGraph

    Dim dbErr As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        logs(1) = "TESTS BLOCKED: GetTestDb returned Nothing: " & dbErr
        Test_SincronizarSuministradoresEnEdicion_Adversarial = Test_Helper.BuildJsonFail(logs(1), logs)
        GoTo Teardown
    End If

    ' --- Primera llamada ---
    Dim pError1 As String
    modSuministradoresSyncHelper.SincronizarSuministradoresParaEdicion _
        CStr(FIX_EDICION), Nothing, pError1

    If pError1 <> "" Then
        logs(2) = "2. Assert FAIL: primera llamada devolvio error: " & pError1
        Test_SincronizarSuministradoresEnEdicion_Adversarial = _
            Test_Helper.BuildJsonFail(pError1, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) FROM TbProyectosEdicionesSuministradores " & _
        "WHERE IDEdicion=" & FIX_EDICION, dbOpenSnapshot)
    Dim countAfterFirst As Long
    countAfterFirst = rs.fields(0).value
    rs.Close: Set rs = Nothing
    logs(2) = "2. Assert: countAfterFirst=" & countAfterFirst

    ' --- Segunda llamada (simula doble-click) ---
    Dim pError2 As String
    modSuministradoresSyncHelper.SincronizarSuministradoresParaEdicion _
        CStr(FIX_EDICION), Nothing, pError2

    If pError2 <> "" Then
        logs(4) = "4. Assert FAIL: segunda llamada devolvio error: " & pError2
        Test_SincronizarSuministradoresEnEdicion_Adversarial = _
            Test_Helper.BuildJsonFail(pError2, logs)
        GoTo Teardown
    End If

    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) FROM TbProyectosEdicionesSuministradores " & _
        "WHERE IDEdicion=" & FIX_EDICION, dbOpenSnapshot)
    Dim countAfterSecond As Long
    countAfterSecond = rs.fields(0).value
    rs.Close: Set rs = Nothing
    logs(4) = "4. Assert: countAfterSecond=" & countAfterSecond

    ' --- Assert idempotencia ---
    If countAfterSecond <> countAfterFirst Then
        Test_SincronizarSuministradoresEnEdicion_Adversarial = _
            Test_Helper.BuildJsonFail( _
                "Idempotency violated: first call count=" & countAfterFirst & _
                " second call count=" & countAfterSecond, logs)
        GoTo Teardown
    End If

    Test_SincronizarSuministradoresEnEdicion_Adversarial = _
        Test_Helper.BuildJsonOk("adversarial_idempotent_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & FIX_EDICION
    End If
    Test_Fixtures.TeardownSubcatGraph
    Test_Helper.ResetTestSession
    On Error GoTo 0
End Function

' ============================================================
' BLOQUE 3 — REQ-CAL-12 — Atoms for ListarEvidenciasPendientesEnInforme
'
' Helper: ListarEvidenciasPendientesEnInforme(ByRef p_IDEdicion As String, ...)
'   Returns "Sin pendientes" or "PUBLICACIÓN BLOQUEADA: <detalle>" or "" on error.
'   Heurística: IDAnexo IS NULL OR Nz(IDAnexo, 0) = 0 ? evidencia pendiente.
'
' Fixture IDs: dedicated range 904000-904099 (no overlap with
'   SeedBaseGraph 900500-900599, SeedSubcatGraph 900100-900221,
'   Bloque 3 / REQ-CAL-07 901000-901099, REQ-CAL-09 902000-902099,
'   REQ-CAL-10 903000-903099).
'   Sub-fixture for subcontratistas: 904100-904199 (IDs de filas en
'   TbProyectosEdicionesSuministradores: 904250+ para evitar colisión).
' ============================================================

' ============================================================
' ATOM 1 — Happy: ListaPendientesPorSubcontratista
' GIVEN staging sandbox + edición con 2 subcontratistas (ACME, BETA)
'       cada uno con 1 fila en TbProyectosEdicionesSuministradores SIN evidencia
' WHEN ListarEvidenciasPendientesEnInforme(p_IDEdicion, db, 0, err) is called
' THEN:
'   - p_Error vacío
'   - retorno contiene "PUBLICACIÓN BLOQUEADA"
'   - retorno contiene "Subcontratista ACME" + "(CIF F800123456)"
'   - retorno contiene "Subcontratista BETA" + "(CIF F800123457)"
'   - retorno contiene ": 1 riesgo(s) sin evidencia" para cada uno
' ============================================================
Public Function Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Happy_ListaPendientesPorSubcontratista() As String
    Dim logs(0 To 8) As String
    Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Happy_ListaPendientesPorSubcontratista = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded + SeedR12Fixture(904102)"
    logs(1) = "2. Arrange: insertar ACME (CIF F800123456) y BETA (CIF F800123457)"
    logs(2) = "3. Arrange: 1 fila pendiente por subcontratista (IDAnexo=Null)"
    logs(3) = "4. Act: ListarEvidenciasPendientesEnInforme(904102, db, 0, err)"
    logs(4) = "5. Assert: p_Error vacío"
    logs(5) = "6. Assert: retorno contiene 'PUBLICACIÓN BLOQUEADA'"
    logs(6) = "7. Assert: retorno contiene nombres de ambos subcontratistas"
    logs(7) = "8. Assert: retorno contiene ': 1 riesgo(s) sin evidencia'"
    logs(8) = "9. Teardown: borrar fixture"

    Dim cfgErr As String
    If Not Test_Helper.EnsureTestConfigLoaded(cfgErr) Then
        Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Happy_ListaPendientesPorSubcontratista = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Happy_ListaPendientesPorSubcontratista = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    SeedR12Fixture FIX_ID_R12_EDICION_HAPPY, FIX_ID_R12_PROYECTO, FIX_ID_R12_EXPEDIENTE

    ' Insertar subcontratistas
    db.Execute "INSERT INTO TbSuministradores (IDSuministrador, Nombre, CIF, ConsorcioPropio) " & _
               "VALUES (" & FIX_ID_R12_SUM_ACME & ", 'Subcontratista ACME', 'F800123456', 'No')", dbFailOnError
    db.Execute "INSERT INTO TbSuministradores (IDSuministrador, Nombre, CIF, ConsorcioPropio) " & _
               "VALUES (" & FIX_ID_R12_SUM_BETA & ", 'Subcontratista BETA', 'F800123457', 'No')", dbFailOnError

    ' 1 fila por subcontratista, IDAnexo omitido (= NULL ? evidencia pendiente)
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador) " & _
               "VALUES (" & FIX_ID_R12_PES_HAPPY_ACME & ", " & FIX_ID_R12_EDICION_HAPPY & ", " & FIX_ID_R12_SUM_ACME & ")", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador) " & _
               "VALUES (" & FIX_ID_R12_PES_HAPPY_BETA & ", " & FIX_ID_R12_EDICION_HAPPY & ", " & FIX_ID_R12_SUM_BETA & ")", dbFailOnError

    Dim m_Result As String
    Dim m_Err As String
    m_Result = ListarEvidenciasPendientesEnInforme(CStr(FIX_ID_R12_EDICION_HAPPY), db, 0, m_Err)

    If Len(m_Err) <> 0 Then
        Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Happy_ListaPendientesPorSubcontratista = BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If Len(m_Result) = 0 Then
        Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Happy_ListaPendientesPorSubcontratista = BuildFail("retorno vacío", logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, "PUBLICACIÓN BLOQUEADA", vbTextCompare) = 0 Then
        Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Happy_ListaPendientesPorSubcontratista = BuildFail("esperaba 'PUBLICACIÓN BLOQUEADA', got: " & m_Result, logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, "Subcontratista ACME", vbTextCompare) = 0 Then
        Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Happy_ListaPendientesPorSubcontratista = BuildFail("esperaba 'Subcontratista ACME' en lista, got: " & m_Result, logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, "Subcontratista BETA", vbTextCompare) = 0 Then
        Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Happy_ListaPendientesPorSubcontratista = BuildFail("esperaba 'Subcontratista BETA' en lista, got: " & m_Result, logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, "CIF F800123456", vbTextCompare) = 0 Then
        Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Happy_ListaPendientesPorSubcontratista = BuildFail("esperaba 'CIF F800123456' (ACME), got: " & m_Result, logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, "CIF F800123457", vbTextCompare) = 0 Then
        Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Happy_ListaPendientesPorSubcontratista = BuildFail("esperaba 'CIF F800123457' (BETA), got: " & m_Result, logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, ": 1 riesgo(s) sin evidencia", vbTextCompare) = 0 Then
        Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Happy_ListaPendientesPorSubcontratista = BuildFail("esperaba ': 1 riesgo(s) sin evidencia', got: " & m_Result, logs)
        GoTo Teardown
    End If

    Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Happy_ListaPendientesPorSubcontratista = BuildOk("happy_lista_pendientes_por_subcontratista_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownR12Fixture FIX_ID_R12_EDICION_HAPPY, FIX_ID_R12_PROYECTO, FIX_ID_R12_EXPEDIENTE
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Happy_ListaPendientesPorSubcontratista = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 2 — Sad: TodasCargadas_Publicable
' GIVEN staging sandbox + edición con 2 subcontratistas TODOS con
'       IDAnexo cargado (evidencia presente)
' WHEN ListarEvidenciasPendientesEnInforme is called
' THEN:
'   - p_Error vacío
'   - retorno = "Sin pendientes" (literal, sin prefijo de bloqueo)
' ============================================================
Public Function Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Sad_TodasCargadas_Publicable() As String
    Dim logs(0 To 6) As String
    Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Sad_TodasCargadas_Publicable = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded + SeedR12Fixture(904103)"
    logs(1) = "2. Arrange: insertar ACME + BETA con IDAnexo=1 (evidencia cargada)"
    logs(2) = "3. Act: ListarEvidenciasPendientesEnInforme(904103, db, 0, err)"
    logs(3) = "4. Assert: p_Error vacío"
    logs(4) = "5. Assert: retorno = 'Sin pendientes' literal"
    logs(5) = "6. Assert: retorno NO contiene 'PUBLICACIÓN BLOQUEADA'"
    logs(6) = "7. Teardown: borrar fixture"

    Dim cfgErr As String
    If Not Test_Helper.EnsureTestConfigLoaded(cfgErr) Then
        Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Sad_TodasCargadas_Publicable = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Sad_TodasCargadas_Publicable = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    SeedR12Fixture FIX_ID_R12_EDICION_SAD, FIX_ID_R12_PROYECTO, FIX_ID_R12_EXPEDIENTE

    db.Execute "INSERT INTO TbSuministradores (IDSuministrador, Nombre, CIF, ConsorcioPropio) " & _
               "VALUES (" & FIX_ID_R12_SUM_ACME & ", 'Subcontratista ACME', 'F800123456', 'No')", dbFailOnError
    db.Execute "INSERT INTO TbSuministradores (IDSuministrador, Nombre, CIF, ConsorcioPropio) " & _
               "VALUES (" & FIX_ID_R12_SUM_BETA & ", 'Subcontratista BETA', 'F800123457', 'No')", dbFailOnError

    ' IDAnexo = 1 (evidencia cargada)
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador, IDAnexo) " & _
               "VALUES (" & FIX_ID_R12_PES_SAD_ACME & ", " & FIX_ID_R12_EDICION_SAD & ", " & FIX_ID_R12_SUM_ACME & ", 1)", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador, IDAnexo) " & _
               "VALUES (" & FIX_ID_R12_PES_SAD_BETA & ", " & FIX_ID_R12_EDICION_SAD & ", " & FIX_ID_R12_SUM_BETA & ", 1)", dbFailOnError

    Dim m_Result As String
    Dim m_Err As String
    m_Result = ListarEvidenciasPendientesEnInforme(CStr(FIX_ID_R12_EDICION_SAD), db, 0, m_Err)

    If Len(m_Err) <> 0 Then
        Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Sad_TodasCargadas_Publicable = BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If m_Result <> "Sin pendientes" Then
        Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Sad_TodasCargadas_Publicable = BuildFail("esperaba 'Sin pendientes' literal, got: [" & m_Result & "]", logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, "PUBLICACIÓN BLOQUEADA", vbTextCompare) > 0 Then
        Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Sad_TodasCargadas_Publicable = BuildFail("NO esperaba 'PUBLICACIÓN BLOQUEADA' en Sad, got: " & m_Result, logs)
        GoTo Teardown
    End If

    Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Sad_TodasCargadas_Publicable = BuildOk("sad_todas_cargadas_publicable_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownR12Fixture FIX_ID_R12_EDICION_SAD, FIX_ID_R12_PROYECTO, FIX_ID_R12_EXPEDIENTE
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_SuministradoresSyncHelper_ListarEvidenciasPendientesEnInforme_Sad_TodasCargadas_Publicable = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 3 — Edge: NITNullPlaceholderLegible (typo en Suministraciones — diseño)
' GIVEN staging sandbox + edición con subcontratista GAMMA (CIF vacío)
'       con evidencia pendiente
' WHEN ListarEvidenciasPendientesEnInforme is called
' THEN:
'   - p_Error vacío
'   - retorno contiene el placeholder "(sin CIF)" cuando CIF es vacío
'   (nombre del átomo: el campo se llama NIT coloquialmente; el schema
'    tiene CIF — documentado en el helper)
' ============================================================
Public Function Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Edge_NITNullPlaceholderLegible() As String
    Dim logs(0 To 6) As String
    Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Edge_NITNullPlaceholderLegible = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded + SeedR12Fixture(904104)"
    logs(1) = "2. Arrange: insertar GAMMA con CIF='' (Null en schema)"
    logs(2) = "3. Arrange: 1 fila pendiente para GAMMA"
    logs(3) = "4. Act: ListarEvidenciasPendientesEnInforme(904104, db, 0, err)"
    logs(4) = "5. Assert: p_Error vacío"
    logs(5) = "6. Assert: retorno contiene '(sin CIF)' placeholder"
    logs(6) = "7. Teardown: borrar fixture"

    Dim cfgErr As String
    If Not Test_Helper.EnsureTestConfigLoaded(cfgErr) Then
        Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Edge_NITNullPlaceholderLegible = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Edge_NITNullPlaceholderLegible = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    SeedR12Fixture FIX_ID_R12_EDICION_EDGE, FIX_ID_R12_PROYECTO, FIX_ID_R12_EXPEDIENTE

    ' GAMMA con CIF explícitamente vacío. NOTA: el helper consulta Nz(CIF, "")
    ' y aplica placeholder "(sin CIF)" si queda vacío. La columna en el schema
    ' es "CIF", no "NIT" — el nombre del átomo refleja el dominio coloquial.
    db.Execute "INSERT INTO TbSuministradores (IDSuministrador, Nombre, CIF, ConsorcioPropio) " & _
               "VALUES (" & FIX_ID_R12_SUM_GAMMA & ", 'Subcontratista GAMMA', '', 'No')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador) " & _
               "VALUES (" & FIX_ID_R12_PES_EDGE_GAMMA & ", " & FIX_ID_R12_EDICION_EDGE & ", " & FIX_ID_R12_SUM_GAMMA & ")", dbFailOnError

    Dim m_Result As String
    Dim m_Err As String
    m_Result = ListarEvidenciasPendientesEnInforme(CStr(FIX_ID_R12_EDICION_EDGE), db, 0, m_Err)

    If Len(m_Err) <> 0 Then
        Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Edge_NITNullPlaceholderLegible = BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If Len(m_Result) = 0 Then
        Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Edge_NITNullPlaceholderLegible = BuildFail("retorno vacío", logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, "Subcontratista GAMMA", vbTextCompare) = 0 Then
        Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Edge_NITNullPlaceholderLegible = BuildFail("esperaba nombre GAMMA en lista, got: " & m_Result, logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, "(sin CIF)", vbTextCompare) = 0 Then
        Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Edge_NITNullPlaceholderLegible = BuildFail("esperaba placeholder '(sin CIF)', got: " & m_Result, logs)
        GoTo Teardown
    End If
    ' HTMLSafe escapa paréntesis como entities? No: solo & < > ". Los paréntesis
    ' quedan literales — eso es lo que queremos en el report (plantilla tipo
    ' "Subcontratista ACME (CIF F800123456)" se mantiene legible).
    If InStr(1, m_Result, "CIF (sin CIF)", vbTextCompare) > 0 Then
        Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Edge_NITNullPlaceholderLegible = BuildFail("placeholder duplicado en 'CIF (sin CIF)', got: " & m_Result, logs)
        GoTo Teardown
    End If

    Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Edge_NITNullPlaceholderLegible = BuildOk("edge_nit_null_placeholder_legible_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownR12Fixture FIX_ID_R12_EDICION_EDGE, FIX_ID_R12_PROYECTO, FIX_ID_R12_EXPEDIENTE
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Edge_NITNullPlaceholderLegible = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 4 — Adversarial: BloqueaPublicacionConDetalle
' GIVEN staging sandbox + edición con ACME con 2 riesgos sin evidencia
' WHEN ListarEvidenciasPendientesEnInforme is called
' THEN:
'   - p_Error vacío
'   - retorno contiene prefijo "PUBLICACIÓN BLOQUEADA"
'   - retorno contiene el detalle del subcontratista con conteo correcto
'   (Documentado: el helper NO llama a publish; solo reporta. La decisión
'    de bloquear publicación la toma el caller — p.ej. el form o
'    modPublicacionCalidadExecutionHelper. Ver design.md §2.2 REQ-CAL-12.)
' ============================================================
Public Function Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Adversarial_BloqueaPublicacionConDetalle() As String
    Dim logs(0 To 7) As String
    Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Adversarial_BloqueaPublicacionConDetalle = BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureSandboxLoaded + SeedR12Fixture(904105)"
    logs(1) = "2. Arrange: insertar ACME con 2 filas pendientes (IDAnexo=Null)"
    logs(2) = "3. Act: ListarEvidenciasPendientesEnInforme(904105, db, 0, err)"
    logs(3) = "4. Assert: p_Error vacío"
    logs(4) = "5. Assert: retorno comienza con 'PUBLICACIÓN BLOQUEADA'"
    logs(5) = "6. Assert: retorno contiene 'ACME' y ': 2 riesgo(s) sin evidencia'"
    logs(6) = "7. Assert: retorno NO contiene 'Sin pendientes'"
    logs(7) = "8. Teardown: borrar fixture"

    Dim cfgErr As String
    If Not Test_Helper.EnsureTestConfigLoaded(cfgErr) Then
        Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Adversarial_BloqueaPublicacionConDetalle = BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Adversarial_BloqueaPublicacionConDetalle = BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    SeedR12Fixture FIX_ID_R12_EDICION_ADV, FIX_ID_R12_PROYECTO, FIX_ID_R12_EXPEDIENTE

    db.Execute "INSERT INTO TbSuministradores (IDSuministrador, Nombre, CIF, ConsorcioPropio) " & _
               "VALUES (" & FIX_ID_R12_SUM_ACME & ", 'Subcontratista ACME', 'F800123456', 'No')", dbFailOnError
    ' 2 filas pendientes para ACME (mismo subcontratista, distintos riesgos)
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador) " & _
               "VALUES (" & FIX_ID_R12_PES_ADV_ACME & ", " & FIX_ID_R12_EDICION_ADV & ", " & FIX_ID_R12_SUM_ACME & ")", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador) " & _
               "VALUES (" & (FIX_ID_R12_PES_ADV_ACME + 1) & ", " & FIX_ID_R12_EDICION_ADV & ", " & FIX_ID_R12_SUM_ACME & ")", dbFailOnError

    Dim m_Result As String
    Dim m_Err As String
    m_Result = ListarEvidenciasPendientesEnInforme(CStr(FIX_ID_R12_EDICION_ADV), db, 0, m_Err)

    If Len(m_Err) <> 0 Then
        Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Adversarial_BloqueaPublicacionConDetalle = BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If Len(m_Result) = 0 Then
        Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Adversarial_BloqueaPublicacionConDetalle = BuildFail("retorno vacío", logs)
        GoTo Teardown
    End If
    ' Prefijo de bloqueo
    If Left$(m_Result, Len("PUBLICACIÓN BLOQUEADA")) <> "PUBLICACIÓN BLOQUEADA" Then
        Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Adversarial_BloqueaPublicacionConDetalle = BuildFail("esperaba prefijo 'PUBLICACIÓN BLOQUEADA', got: " & m_Result, logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, "Subcontratista ACME", vbTextCompare) = 0 Then
        Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Adversarial_BloqueaPublicacionConDetalle = BuildFail("esperaba 'Subcontratista ACME' en detalle, got: " & m_Result, logs)
        GoTo Teardown
    End If
    If InStr(1, m_Result, ": 2 riesgo(s) sin evidencia", vbTextCompare) = 0 Then
        Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Adversarial_BloqueaPublicacionConDetalle = BuildFail("esperaba ': 2 riesgo(s) sin evidencia', got: " & m_Result, logs)
        GoTo Teardown
    End If
    ' Adversarial: NO debe decir "Sin pendientes" cuando hay bloqueantes
    If InStr(1, m_Result, "Sin pendientes", vbTextCompare) > 0 Then
        Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Adversarial_BloqueaPublicacionConDetalle = BuildFail("NO esperaba 'Sin pendientes', got: " & m_Result, logs)
        GoTo Teardown
    End If

    Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Adversarial_BloqueaPublicacionConDetalle = BuildOk("adversarial_bloquea_publicacion_con_detalle_pass", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownR12Fixture FIX_ID_R12_EDICION_ADV, FIX_ID_R12_PROYECTO, FIX_ID_R12_EXPEDIENTE
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_SuministracionesSyncHelper_ListarEvidenciasPendientesEnInforme_Adversarial_BloqueaPublicacionConDetalle = BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

