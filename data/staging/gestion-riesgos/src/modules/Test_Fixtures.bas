Attribute VB_Name = "Test_Fixtures"
Option Compare Database
Option Explicit

' ============================================================
' Test_Fixtures — Grafo de fixtures para GESTIÓN_RIESGOS
'
' Skill: access-vba-tdd v1.5 §3.4 + §3.5
'
' PATRÓN DE DOS CAPAS:
'
'  TIER 1 — Grafo base (SeedAll / TeardownAll):
'    Se ejecuta UNA VEZ por suite desde RunAll.
'    Inserta TODAS las entidades padre necesarias.
'    TeardownAll en orden inverso (hijos ? padres).
'
'  TIER 2 — Delta por test:
'    Cada test inserta solo su delta sobre el grafo base.
'    Teardown_Item solo borra su delta, no el grafo base.
'
' Fixture IDs en rango = 900000 (FIX_ID_BASE).
' ============================================================

Private Const FIX_ID_BASE As Long = 900000

' --- Cache de conexión (v1.9 §2) ---
Private m_TestDb As DAO.Database

' --- IDs del grafo base (Tier 1) ---
Private m_idExpediente  As Long
Private m_idProyecto   As Long
Private m_idEdicion     As Long
Private m_idRiesgo      As Long
Private m_idPM          As Long
Private m_idPC          As Long

' --- IDs del grafo Subcontratistas (Tier 1) ---
Private m_idExpediente_Subcat As Long
Private m_idProyecto_Subcat As Long
Private m_idEdicion_Subcat As Long

' --- Getters ---
Public Property Get Cache_ExpedienteId() As Long: Cache_ExpedienteId = m_idExpediente: End Property
Public Property Get Cache_ProyectoId() As Long: Cache_ProyectoId = m_idProyecto: End Property
Public Property Get Cache_EdicionId() As Long: Cache_EdicionId = m_idEdicion: End Property
Public Property Get Cache_RiesgoId() As Long: Cache_RiesgoId = m_idRiesgo: End Property
Public Property Get Cache_PMId() As Long: Cache_PMId = m_idPM: End Property
Public Property Get Cache_PCId() As Long: Cache_PCId = m_idPC: End Property
Public Property Get Subcat_EdicionId() As Long: Subcat_EdicionId = m_idEdicion_Subcat: End Property

' ============================================================
' GetTestDb — conexión directa al backend local
' Devuelve DAO.Database o Nothing si falla.
' El caller debe chequear con If db Is Nothing Then ...
' Los errores se devuelven como JSON en p_Error, nunca como Err.Raise.
' ============================================================
Public Function GetTestDb(Optional ByRef p_Error As String) As DAO.Database
    Set GetTestDb = Nothing
    p_Error = ""

    ' v1.9 §2: reuse cached connection if alive
    If Not m_TestDb Is Nothing Then
        On Error Resume Next
        Dim dummyName As String
        dummyName = m_TestDb.Name  ' ? validate connection still open
        If Err.Number = 0 Then
            Set GetTestDb = m_TestDb
            Exit Function
        End If
        ' Cache corrupt (Access closed it) — fall through to reopen
        On Error GoTo 0
        Set m_TestDb = Nothing
    End If

    Dim cfgError As String
    If Not ForceLocalBackend(cfgError) Then
        p_Error = "{""ok"":false,""error"":""TESTS BLOCKED: ForceLocalBackend failed: " & EscapeJsonString(cfgError) & """,""logs"":[]}"
        Exit Function
    End If

    Dim ws As DAO.Workspace
    Dim dbPath As String
    Dim localPassword As String

    dbPath = m_BackendSandboxURL
    If dbPath = "" Or Not CreateObject("Scripting.FileSystemObject").FileExists(dbPath) Then
        p_Error = "{""ok"":false,""error"":""TESTS BLOCKED: BackendSandbox not found: " & EscapeJsonString(dbPath) & """,""logs"":[]}"
        Exit Function
    End If

    localPassword = m_PasswordBackend
    Set ws = DBEngine(0)
    On Error Resume Next
    Set m_TestDb = ws.OpenDatabase(dbPath, False, False, ";PWD=" & localPassword)
    If Err.Number <> 0 Then
        p_Error = "{""ok"":false,""error"":""TESTS BLOCKED: Cannot open backend: " & EscapeJsonString(Err.Description) & """,""logs"":[]}"
        Set m_TestDb = Nothing
        Exit Function
    End If
    On Error GoTo 0
    Set GetTestDb = m_TestDb
End Function

Private Function GetNextId(table As String, idCol As String) As Long
    Dim rs As DAO.Recordset
    Dim maxId As Long
    maxId = FIX_ID_BASE
    On Error Resume Next
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        On Error GoTo 0
        GetNextId = maxId + 10
        Exit Function
    End If
    Set rs = db.OpenRecordset("SELECT MAX(" & idCol & ") FROM " & table)
    If Not rs.EOF And Not IsNull(rs.Fields(0).value) Then
        maxId = rs.Fields(0).value
    End If
    rs.Close: Set rs = Nothing
    Set db = Nothing
    On Error GoTo 0
    GetNextId = maxId + 10
End Function

Private Function IsFixtureId(id As Long) As Boolean
    IsFixtureId = (id >= FIX_ID_BASE)
End Function

Private Function SafeClng(id As String) As Long
    On Error Resume Next
    SafeClng = CLng(Nz(id, 0))
    If Err.Number <> 0 Then
        Err.Clear
        SafeClng = 0
    End If
    On Error GoTo 0
End Function

' ============================================================
' TIER 1 — Grafo Base
'
' SeedAll: limpia rango FIX_ID_BASE e inserta grafo completo.
' TeardownAll: borra en orden inverso (hijos ? padres).
' ============================================================

Public Sub SeedAll()
    On Error GoTo EH_SeedAll
    SeedBaseGraph
    SeedSubcatGraph
    Exit Sub

EH_SeedAll:
    Err.Raise Err.Number, "Test_Fixtures.SeedAll", "SeedAll failed: " & Err.Description
End Sub

Public Sub TeardownAll()
    On Error Resume Next
    ' Orden inverso: hijos antes que padres
    TeardownSubcatGraph
    TeardownBaseGraph
    CloseTestDb  ' v1.9 §2 SuiteTeardown
    Test_Helper.ResetTestSession  ' v2.0 §3: reset EVE flag para próxima sesión
    On Error GoTo 0
End Sub

' v1.9 §2: wrapper that resets local m_TestDb after calling Test_Helper.CloseTestDb
Public Sub CloseTestDb()
    On Error Resume Next
    Test_Helper.CloseTestDb
    Set m_TestDb = Nothing
    On Error GoTo 0
End Sub

' -- Base Graph ------------------------------------------------
' Fully self-contained fixture: creates its own Expediente, Proyecto,
' Edición, Riesgo, PM, PC — all in fixture ID range. No external data.
' ============================================================

Private Sub SeedBaseGraph()
    On Error GoTo EH_SeedBase
    Dim dbErr As String
    Dim db As DAO.Database
    Dim e As Long
    Dim d As String
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Err.Raise 1001, "SeedBaseGraph", "GetTestDb returned Nothing. dbErr=" & dbErr
        Exit Sub
    End If

    ' SeedSubcatGraph uses hardcoded fixture IDs in range 900100-900121.
    ' SeedBaseGraph uses a DIFFERENT non-overlapping range: 900500-900599.

    ' Idempotent cleanup must delete children before parents to respect FK order.
    m_idExpediente = 900500
    m_idProyecto = 900501
    m_idEdicion = 900502
    m_idRiesgo = 900503
    m_idPM = 900504
    m_idPC = 900505

    db.Execute "DELETE FROM TbRiesgosPlanContingenciaDetalle WHERE IDContingencia=" & m_idPC
    db.Execute "DELETE FROM TbRiesgosPlanMitigacionDetalle WHERE IDMitigacion=" & m_idPM
    db.Execute "DELETE FROM TbRiesgosPlanContingenciaPpal WHERE IDContingencia=" & m_idPC
    db.Execute "DELETE FROM TbRiesgosPlanMitigacionPpal WHERE IDMitigacion=" & m_idPM
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & m_idRiesgo
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & m_idEdicion
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & m_idProyecto
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & m_idExpediente

    ' 1. TbExpedientes (padre de TbProyectos) — range 900500+
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
        "VALUES (" & m_idExpediente & ", 'TESTFIX', 'Fixture expediente test', 'Test', 1)"

    ' 2. TbProyectos (hijo de TbExpedientes) — range 900501
    db.Execute "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto) " & _
        "VALUES (" & m_idProyecto & ", " & m_idExpediente & ", 'TESTPROJ" & m_idProyecto & "')"

    ' 3. TbProyectosEdiciones (hijo de TbProyectos) — range 900502
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
        "VALUES (" & m_idEdicion & ", " & m_idProyecto & ", 1, 'TESTUSER')"

    ' 4. TbRiesgos (hijo de TbProyectosEdiciones) — range 900503
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoRiesgo, Descripcion, Estado, Priorizacion, " & _
        "CodigoUnico, FechaDetectado, DetectadoPor, EntidadDetecta, Plazo, Calidad, Coste, ImpactoGlobal, " & _
        "Vulnerabilidad, Valoracion, Mitigacion, Contingencia, RequierePlanContingencia) " & _
        "VALUES (" & m_idRiesgo & ", " & m_idEdicion & ", 'TST" & Right$(CStr(m_idRiesgo), 4) & "', " & _
        "'Fixture riesgo test', 'Detectado', 3, " & _
        "'UNICO-TEST-" & m_idRiesgo & "', #" & Format$(Now, "yyyy-mm-dd") & "#, " & _
        "'TESTUSER', 'TEST', 'Medio', 'Medio', 'Medio', 'Medio', 'Medio', 'Medio', 'Reducir', 'Sí', 'Sí')"

    ' 5. TbRiesgosPlanMitigacionPpal (hijo de TbRiesgos) — range 900504
    db.Execute "INSERT INTO TbRiesgosPlanMitigacionPpal (IDMitigacion, IDRiesgo, CodMitigacion, DisparadorDelPlan, Estado) " & _
        "VALUES (" & m_idPM & ", " & m_idRiesgo & ", 'PMTST" & Right$(CStr(m_idPM), 4) & "', " & _
        "'Trigger fixture PM test', 'Definido')"

    ' 6. TbRiesgosPlanContingenciaPpal (hijo de TbRiesgos) — range 900505
    db.Execute "INSERT INTO TbRiesgosPlanContingenciaPpal (IDContingencia, IDRiesgo, CodContingencia, DisparadorDelPlan, Estado) " & _
        "VALUES (" & m_idPC & ", " & m_idRiesgo & ", 'PCTST" & Right$(CStr(m_idPC), 4) & "', " & _
        "'Trigger fixture PC test', 'Definido')"

    Set db = Nothing
    Exit Sub

EH_SeedBase:
    e = Err.Number
    d = Err.Description
    On Error Resume Next
    If Not db Is Nothing Then Set db = Nothing
    Err.Raise e, "SeedBaseGraph", "SeedBaseGraph failed: " & e & " - " & d
End Sub

Private Sub TeardownBaseGraph()
    On Error Resume Next
    Dim dbErr As String
    Dim db As DAO.Database
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then Set db = Nothing: Exit Sub

    If IsFixtureId(m_idPC) Then db.Execute "DELETE FROM TbRiesgosPlanContingenciaDetalle WHERE IDContingencia=" & m_idPC
    If IsFixtureId(m_idPM) Then db.Execute "DELETE FROM TbRiesgosPlanMitigacionDetalle WHERE IDMitigacion=" & m_idPM
    If IsFixtureId(m_idPC) Then db.Execute "DELETE FROM TbRiesgosPlanContingenciaPpal WHERE IDContingencia=" & m_idPC
    If IsFixtureId(m_idPM) Then db.Execute "DELETE FROM TbRiesgosPlanMitigacionPpal WHERE IDMitigacion=" & m_idPM
    If IsFixtureId(m_idRiesgo) Then db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & m_idRiesgo
    If IsFixtureId(m_idEdicion) Then db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & m_idEdicion
    If IsFixtureId(m_idProyecto) Then db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & m_idProyecto
    If IsFixtureId(m_idExpediente) Then db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & m_idExpediente
    Set db = Nothing
    On Error GoTo 0
End Sub

' -- Subcontratistas Graph ------------------------------------
' Creates its OWN complete fixture hierarchy:
'   - Expediente fixture (900xxx)
'   - Proyecto hijo del expediente fixture
'   - Edición hija del proyecto fixture
'   - TbExpedientesSuministradores: hierarchy of suppliers with ConsorcioPropio='Sí'
'   - TbProyectosEdicionesSuministradores: suppliers in the edition
'
' NO dependency on existing expediente 1004 or any real data.
' Each test runs on its own fixture.
' ============================================================

Public Sub SeedSubcatGraph()
    On Error GoTo EH_SeedSubcat
    Dim dbErr As String
    Dim db As DAO.Database
    Const FIX_EXPEDIENTE As Long = 900100
    Const FIX_PROYECTO   As Long = 900101
    Const FIX_EDICION    As Long = 900102
    Const FIX_ROOT       As Long = 900110
    Const FIX_CHILD1     As Long = 900120
    Const FIX_CHILD2     As Long = 900130
    Const FIX_GRANDCHILD As Long = 900121
    Const FIX_IDES_ROOT   As Long = 900200
    Const FIX_IDES_CHILD1 As Long = 900210
    Const FIX_IDES_CHILD2 As Long = 900220
    Const FIX_IDES_GCHILD As Long = 900211

    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Err.Raise 1001, "SeedSubcatGraph", "GetTestDb returned Nothing. dbErr=" & dbErr
        Exit Sub
    End If

    ' Hardcoded fixture IDs — NUNCA GetNextId para evitar colisiones con datos reales.
    ' Rango usado: 900100-900121 (separado de SeedBaseGraph 900500-900599).
    m_idExpediente_Subcat = FIX_EXPEDIENTE
    m_idProyecto_Subcat = FIX_PROYECTO
    m_idEdicion_Subcat = FIX_EDICION
    TeardownSubcatGraph

    ' 1. Expediente fixture
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & m_idExpediente_Subcat
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
        "VALUES (" & m_idExpediente_Subcat & ", 'TESTSUB', 'Fixture subcontratistas test', 'Test', 1)"

    ' 2. Proyecto fixture
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & m_idProyecto_Subcat
    db.Execute "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto) " & _
        "VALUES (" & m_idProyecto_Subcat & ", " & m_idExpediente_Subcat & ", 'TESTSUBPROJ')"

    ' 3. Edición fixture
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & m_idEdicion_Subcat
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
        "VALUES (" & m_idEdicion_Subcat & ", " & m_idProyecto_Subcat & ", 99, 'TESTUSER')"

    ' 4. Limpiar suppliers de la edición fixture
    db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & m_idEdicion_Subcat

    ' 5. Poblar TbSuministradores con jerarquía de test
    db.Execute "DELETE FROM TbSuministradores WHERE IDSuministrador IN (" & FIX_ROOT & "," & FIX_CHILD1 & "," & FIX_CHILD2 & "," & FIX_GRANDCHILD & ")"
    db.Execute "INSERT INTO TbSuministradores (IDSuministrador, Nombre, CIF, ConsorcioPropio) " & _
        "VALUES (" & FIX_ROOT & ", 'FixtureRoot', 'F99999001', 'Sí')"
    db.Execute "INSERT INTO TbSuministradores (IDSuministrador, Nombre, CIF, ConsorcioPropio) " & _
        "VALUES (" & FIX_CHILD1 & ", 'FixtureChild1', 'F99999002', 'Sí')"
    db.Execute "INSERT INTO TbSuministradores (IDSuministrador, Nombre, CIF, ConsorcioPropio) " & _
        "VALUES (" & FIX_CHILD2 & ", 'FixtureChild2', 'F99999003', 'Sí')"
    db.Execute "INSERT INTO TbSuministradores (IDSuministrador, Nombre, CIF, ConsorcioPropio) " & _
        "VALUES (" & FIX_GRANDCHILD & ", 'FixtureGrandChild', 'F99999004', 'Sí')"

    ' 6. Poblar TbExpedientesSuministradores
    db.Execute "DELETE FROM TbExpedientesSuministradores WHERE IDExpediente=" & m_idExpediente_Subcat
    db.Execute "INSERT INTO TbExpedientesSuministradores (IDExpedienteSuministrador, IDExpediente, IDSuministrador, IDPadre) " & _
        "VALUES (" & FIX_IDES_ROOT & ", " & m_idExpediente_Subcat & ", " & FIX_ROOT & ", 0)"
    db.Execute "INSERT INTO TbExpedientesSuministradores (IDExpedienteSuministrador, IDExpediente, IDSuministrador, IDPadre) " & _
        "VALUES (" & FIX_IDES_CHILD1 & ", " & m_idExpediente_Subcat & ", " & FIX_CHILD1 & ", " & FIX_IDES_ROOT & ")"
    db.Execute "INSERT INTO TbExpedientesSuministradores (IDExpedienteSuministrador, IDExpediente, IDSuministrador, IDPadre) " & _
        "VALUES (" & FIX_IDES_CHILD2 & ", " & m_idExpediente_Subcat & ", " & FIX_CHILD2 & ", " & FIX_IDES_ROOT & ")"
    db.Execute "INSERT INTO TbExpedientesSuministradores (IDExpedienteSuministrador, IDExpediente, IDSuministrador, IDPadre) " & _
        "VALUES (" & FIX_IDES_GCHILD & ", " & m_idExpediente_Subcat & ", " & FIX_GRANDCHILD & ", " & FIX_IDES_CHILD1 & ")"

    ' 7. Poblar TbProyectosEdicionesSuministradores
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador) " & _
        "VALUES (" & (FIX_IDES_ROOT + 50) & ", " & m_idEdicion_Subcat & ", " & FIX_CHILD1 & ")"
    db.Execute "INSERT INTO TbProyectosEdicionesSuministradores (ID, IDEdicion, IDSuministrador) " & _
        "VALUES (" & (FIX_IDES_ROOT + 51) & ", " & m_idEdicion_Subcat & ", " & FIX_CHILD2 & ")"

    Set db = Nothing
    Exit Sub

EH_SeedSubcat:
    Dim e2 As Long: e2 = Err.Number
    Dim d2 As String: d2 = Err.Description
    On Error Resume Next
    If Not db Is Nothing Then Set db = Nothing
    Err.Raise e2, "SeedSubcatGraph", "SeedSubcatGraph failed: " & e2 & " - " & d2
End Sub

Public Sub TeardownSubcatGraph()
    Dim db As DAO.Database
    Dim dbErr As String
    Const FIX_ROOT      As Long = 900110
    Const FIX_CHILD1    As Long = 900120
    Const FIX_CHILD2    As Long = 900130
    Const FIX_GRANDCHILD As Long = 900121
    Const FIX_IDES_ROOT   As Long = 900200
    Const FIX_IDES_CHILD1 As Long = 900210
    Const FIX_IDES_CHILD2 As Long = 900220
    Const FIX_IDES_GCHILD As Long = 900211

    If Not IsFixtureId(m_idEdicion_Subcat) And Not IsFixtureId(m_idExpediente_Subcat) Then
        ' Neither fixture was set — nothing to teardown
        Exit Sub
    End If
    On Error Resume Next
    Set db = GetTestDb(dbErr)
    If db Is Nothing Then Set db = Nothing: Exit Sub

    ' Limpiar proveedores de edición
    If IsFixtureId(m_idEdicion_Subcat) Then
        db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion=" & m_idEdicion_Subcat
    End If
    ' Limpiar expediente-proveedores en orden hijo -> padre por la jerarquia IDPadre
    If IsFixtureId(m_idExpediente_Subcat) Then
        db.Execute "DELETE FROM TbExpedientesSuministradores WHERE IDExpedienteSuministrador=" & FIX_IDES_GCHILD
        db.Execute "DELETE FROM TbExpedientesSuministradores WHERE IDExpedienteSuministrador IN (" & FIX_IDES_CHILD1 & "," & FIX_IDES_CHILD2 & ")"
        db.Execute "DELETE FROM TbExpedientesSuministradores WHERE IDExpedienteSuministrador=" & FIX_IDES_ROOT
        db.Execute "DELETE FROM TbExpedientesSuministradores WHERE IDExpediente=" & m_idExpediente_Subcat
    End If
    ' Limpiar proveedores de test
    db.Execute "DELETE FROM TbSuministradores WHERE IDSuministrador IN (" & FIX_ROOT & "," & FIX_CHILD1 & "," & FIX_CHILD2 & "," & FIX_GRANDCHILD & ")"
    ' Limpiar edición, proyecto, expediente (orden inverso)
    If IsFixtureId(m_idEdicion_Subcat) Then
        db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & m_idEdicion_Subcat
    End If
    If IsFixtureId(m_idProyecto_Subcat) Then
        db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & m_idProyecto_Subcat
    End If
    If IsFixtureId(m_idExpediente_Subcat) Then
        db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & m_idExpediente_Subcat
    End If
    Set db = Nothing
    On Error GoTo 0
End Sub

' ============================================================
' TIER 2 — Delta por test
'
' Cada test individual inserta y limpia SOLO su delta.
' NO tocan el grafo base (padres ya existentes).
' ============================================================

' -- Delta: TbRiesgos extra para tests adicionales -------------

Public Function SeedRiesgoExtra(ByRef outId As String) As Boolean
    ' Inserta un riesgo adicional sobre el grafo base ya existente.
    ' El padre (m_idEdicion) ya existe gracias a SeedAll.
    Dim nextId As Long
    nextId = GetNextId("TbRiesgos", "IDRiesgo")
    outId = CStr(nextId)
    GetTestDb().Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & nextId
    GetTestDb.Execute ("INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoRiesgo, Descripcion, Estado, Priorizacion, " & _
        "CodigoUnico, FechaDetectado, DetectadoPor, EntidadDetecta, Plazo, Calidad, Coste, ImpactoGlobal, " & _
        "Vulnerabilidad, Valoracion, Mitigacion, Contingencia, RequierePlanContingencia) " & _
        "VALUES (" & nextId & ", " & m_idEdicion & ", 'TSTX" & Right$(CStr(nextId), 4) & "', " & _
        "'Fixture riesgo extra test', 'Detectado', 3, " & _
        "'UNICO-TEST-X-" & nextId & "', #" & Format$(Now, "yyyy-mm-dd") & "#, " & _
        "'TESTUSER', 'TEST', 'Medio', 'Medio', 'Medio', 'Medio', 'Medio', 'Medio', 'Reducir', 'Sí', 'Sí')")
    SeedRiesgoExtra = True
End Function

Public Sub TeardownRiesgoExtra(id As String)
    Dim numId As Long
    numId = SafeClng(id)
    If IsFixtureId(numId) Then
        On Error Resume Next
        GetTestDb.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo=" & numId
        On Error GoTo 0
    End If
End Sub

Public Function Test_Fixtures_SeedAll_BaseGraphIntegrity() As String
    Dim logs(0 To 6) As String
    Dim dbErr As String
    Dim db As DAO.Database

    logs(0) = "1. Arrange: ForceLocalBackend"
    logs(1) = "2. Act: SeedAll"
    logs(2) = "3. Assert: base graph rows exist"
    logs(3) = "4. Assert: subcat graph rows exist"
    logs(4) = "5. Assert: base IDs were not overwritten by subcat seed"
    logs(5) = "6. Teardown: TeardownAll"
    logs(6) = "7. Assert: seed errors are surfaced as JSON failures"

    On Error GoTo EH

    If Not Test_Helper.ForceLocalBackend(dbErr) Then
        Test_Fixtures_SeedAll_BaseGraphIntegrity = Test_Helper.BuildJsonFail("TESTS BLOCKED: " & dbErr, logs)
        Exit Function
    End If

    SeedAll

    Set db = GetTestDb(dbErr)
    If db Is Nothing Then
        Test_Fixtures_SeedAll_BaseGraphIntegrity = Test_Helper.BuildJsonFail("GetTestDb returned Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    If FixtureCount(db, "TbExpedientes", "IDExpediente=900500") <> 1 Then GoTo BaseGraphFailed
    If FixtureCount(db, "TbProyectos", "IDProyecto=900501") <> 1 Then GoTo BaseGraphFailed
    If FixtureCount(db, "TbProyectosEdiciones", "IDEdicion=900502") <> 1 Then GoTo BaseGraphFailed
    If FixtureCount(db, "TbRiesgos", "IDRiesgo=900503") <> 1 Then GoTo BaseGraphFailed
    If FixtureCount(db, "TbRiesgosPlanMitigacionPpal", "IDMitigacion=900504") <> 1 Then GoTo BaseGraphFailed
    If FixtureCount(db, "TbRiesgosPlanContingenciaPpal", "IDContingencia=900505") <> 1 Then GoTo BaseGraphFailed

    If FixtureCount(db, "TbExpedientes", "IDExpediente=900100") <> 1 Then GoTo SubcatGraphFailed
    If FixtureCount(db, "TbProyectos", "IDProyecto=900101") <> 1 Then GoTo SubcatGraphFailed
    If FixtureCount(db, "TbProyectosEdiciones", "IDEdicion=900102") <> 1 Then GoTo SubcatGraphFailed

    If Cache_ExpedienteId <> 900500 Then GoTo BaseIdsOverwritten
    If Cache_ProyectoId <> 900501 Then GoTo BaseIdsOverwritten
    If Cache_EdicionId <> 900502 Then GoTo BaseIdsOverwritten

    Test_Fixtures_SeedAll_BaseGraphIntegrity = Test_Helper.BuildJsonOk("seedall_integrity_pass", logs)
    GoTo Teardown

BaseGraphFailed:
    Test_Fixtures_SeedAll_BaseGraphIntegrity = Test_Helper.BuildJsonFail("SeedAll did not create the complete base graph", logs)
    GoTo Teardown

SubcatGraphFailed:
    Test_Fixtures_SeedAll_BaseGraphIntegrity = Test_Helper.BuildJsonFail("SeedAll did not create the complete subcat graph", logs)
    GoTo Teardown

BaseIdsOverwritten:
    Test_Fixtures_SeedAll_BaseGraphIntegrity = Test_Helper.BuildJsonFail("SeedSubcatGraph overwrote base fixture cache IDs", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    Set db = Nothing
    TeardownAll
    On Error GoTo 0
    Exit Function

EH:
    Dim errText As String
    errText = Err.Description
    On Error Resume Next
    Set db = Nothing
    TeardownAll
    On Error GoTo 0
    Test_Fixtures_SeedAll_BaseGraphIntegrity = Test_Helper.BuildJsonFail(errText, logs)
End Function

Private Function FixtureCount(ByVal p_db As DAO.Database, ByVal p_Table As String, ByVal p_Where As String) As Long
    Dim rs As DAO.Recordset
    Set rs = p_db.OpenRecordset("SELECT COUNT(*) AS Cnt FROM " & p_Table & " WHERE " & p_Where)
    If Not rs.EOF Then FixtureCount = CLng(Nz(rs.Fields("Cnt").value, 0))
    rs.Close
    Set rs = Nothing
End Function

' ============================================================
' EscapeJsonString — para construir JSON seguro en tests
' ============================================================
Private Function EscapeJsonString(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJsonString = s
End Function

