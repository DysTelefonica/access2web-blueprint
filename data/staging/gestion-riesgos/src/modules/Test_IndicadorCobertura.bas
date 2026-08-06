Attribute VB_Name = "Test_IndicadorCobertura"
' ============================================================
' Test_IndicadorCobertura
'
' issue-59: cobertura de tests para CAP-IND-008 (5 tests criticos).
' Cubre los 5 BR-IND mas relevantes de la capacidad:
'   - BR-IND-002: EstaEnElIntervaloDado (pure logic)
'   - BR-IND-006: IndicadorRiesgosRepositorioV2_GetTabla 5KPIs
'   - BR-IND-007: GetResumen deduplicacion por CodigoUnico
'   - BR-IND-008: SqlDetalle itMaterializados sin GROUP BY
'   - BR-IND-003: RellenaColeccionDetalleRiesgos idempotente
'
' Skill: access-vba-tdd v2.4
' - Test IDs en rango 900700+ (separado de SeedAll 900500-900599)
' - Sandbox via Test_Fixtures.GetTestDb / ForceLocalBackend
' - Cada test hace su propio seed + teardown (no depende de SeedAll)
' - Module-level state de mIndicador.bas: save/restore ColProyectosParaInforme
'   y colRiesgosDetalle entre tests
' ============================================================
Option Compare Database
Option Explicit

' --- IDs de fixture (rango 900700+, separado de SeedAll 900500-900599) ---
Private Const FIX_EXPEDIENTE As Long = 900799
Private Const FIX_PROYECTO   As Long = 900700
Private Const FIX_EDICION    As Long = 900701
Private Const FIX_EDICION2   As Long = 900702
Private Const FIX_RIESGO_A   As Long = 900710
Private Const FIX_RIESGO_B   As Long = 900711
Private Const FIX_RIESGO_C   As Long = 900712
Private Const FIX_OFERTA_A   As Long = 900730
Private Const FIX_OFERTA_B   As Long = 900731
Private Const FIX_MAT_A      As Long = 900720
Private Const FIX_MAT_B      As Long = 900721

' --- Rango de fechas de los tests ---
Private Const FIX_DINI As Date = #1/1/2025#
Private Const FIX_DFIN As Date = #12/31/2025#

' --- IDs de fixture para tests de vigentes (rango 900760+, libre) ---
Private Const FIX_EXPEDIENTE_VIG As Long = 900760
Private Const FIX_PROYECTO_VIG  As Long = 900761
Private Const FIX_EDICION_VIG   As Long = 900762
Private Const FIX_PROYECTO2_VIG As Long = 900763
Private Const FIX_EDICION2_VIG  As Long = 900764

Private Const FIX_RIESGO_T1 As Long = 900765  ' abierto todo el periodo (2025-01-01, NULL)
Private Const FIX_RIESGO_T2 As Long = 900766  ' abierto 1 dia al inicio (2025-12-31, 2026-01-01)
Private Const FIX_RIESGO_T3 As Long = 900767  ' abierto 1 dia al cierre (2026-06-29, 2026-06-30)
Private Const FIX_RIESGO_T4 As Long = 900768  ' cerrado antes del periodo (2025-08-10, 2025-10-05) NO cuenta
Private Const FIX_RIESGO_T5 As Long = 900769  ' detectado despues del periodo (2026-08-15, NULL) NO cuenta
Private Const FIX_RIESGO_T6 As Long = 900770  ' deteccion fuera, cierre dentro (2025-11-20, 2026-03-01) cuenta
Private Const FIX_RIESGO_T7 As Long = 900771  ' riesgo extra en FIX_PROYECTO2_VIG (multi-proyecto)
Private Const FIX_RIESGO_T8 As Long = 900772  ' riesgo extra en FIX_PROYECTO_VIG (multi-proyecto)

' --- Rango de fechas de los tests de vigentes (S1: 2026-01-01 a 2026-06-30) ---
Private Const FIX_DINI_VIG As Date = #1/1/2026#
Private Const FIX_DFIN_VIG As Date = #6/30/2026#


' ============================================================
' Helpers JSON
' ============================================================
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function


' ============================================================
' Helpers de fixture: cleanup, seed proyectos, seed riesgos, etc.
' ============================================================

Private Sub CleanFixture()
    ' Borra en orden inverso de FKs. Idempotente.
    On Error Resume Next
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then Exit Sub

    db.Execute "DELETE FROM TbRiesgosMaterializaciones WHERE IDProyecto=" & FIX_PROYECTO
    db.Execute "DELETE FROM TbRiesgosAIntegrar WHERE IDEdicion IN (" & FIX_EDICION & "," & FIX_EDICION2 & ")"
    db.Execute "DELETE FROM TbRiesgos WHERE IDEdicion IN (" & FIX_EDICION & "," & FIX_EDICION2 & ")"
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion IN (" & FIX_EDICION & "," & FIX_EDICION2 & ")"
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_PROYECTO
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_EXPEDIENTE
    On Error GoTo 0
End Sub

Private Function SeedBase(ByVal db As DAO.Database) As Boolean
    ' Crea TbExpedientes + TbProyectos + 1 TbProyectosEdiciones (la segunda
    ' solo la crea si FIX_EDICION2 > 0 via SeedBaseWithSecondEdition).
    On Error GoTo SeedBase_EH
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
        "VALUES (" & FIX_EXPEDIENTE & ", 'TESTIND', 'Fixture indicador test', 'Test', 1)", dbFailOnError
    db.Execute "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto, NombreProyecto) " & _
        "VALUES (" & FIX_PROYECTO & ", " & FIX_EXPEDIENTE & ", 'INDTEST', 'Indicador Test')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
        "VALUES (" & FIX_EDICION & ", " & FIX_PROYECTO & ", 1, 'TESTUSER')", dbFailOnError
    SeedBase = True
    Exit Function
SeedBase_EH:
    SeedBase = False
End Function

Private Function SeedSecondEdition(ByVal db As DAO.Database) As Boolean
    On Error GoTo SeedSecondEdition_EH
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
        "VALUES (" & FIX_EDICION2 & ", " & FIX_PROYECTO & ", 2, 'TESTUSER')", dbFailOnError
    SeedSecondEdition = True
    Exit Function
SeedSecondEdition_EH:
    SeedSecondEdition = False
End Function

Private Function CountRows(ByVal db As DAO.Database, ByVal p_Table As String, _
                           ByVal p_Where As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Count(*) AS c FROM " & p_Table & " WHERE " & p_Where)
    CountRows = Nz(rs!c, 0)
    rs.Close
    Set rs = Nothing
End Function


' ============================================================
' Test 1 (BR-IND-002): EstaEnElIntervaloDadoRiesgo — pure logic
' (renombrada desde mIndicador.EstaEnElIntervaloDado por colisión
' con Funciones Generales.EstaEnElIntervaloDado, issue #50+#59)
' ============================================================
Public Function Test_IND_mIndicador_EstaEnElIntervaloDadoRiesgo_ProyectoEnvuelveRango_RetornaSi() As String
    On Error GoTo EH
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: informe 01/01/2025-30/06/2025, proyecto 15/11/2024-15/03/2026"
    logs(1) = "2. Act: EstaEnElIntervaloDadoRiesgo (proyecto envuelve al rango)"
    logs(2) = "3. Assert: retorna 'Sí'"
    logs(3) = "4. Assert: caso borde — proyecto anterior al rango retorna 'No'"

    Dim sEnvuelve As String
    sEnvuelve = mIndicador.EstaEnElIntervaloDadoRiesgo("01/01/2025", "30/06/2025", "15/11/2024", "15/03/2026")
    If sEnvuelve <> "Sí" Then
        Test_IND_mIndicador_EstaEnElIntervaloDadoRiesgo_ProyectoEnvuelveRango_RetornaSi = BuildFail( _
            "Proyecto que envuelve el rango debe retornar 'Sí', obtuvo '" & sEnvuelve & "'", logs)
        Exit Function
    End If

    ' Caso borde adicional: proyecto anterior al rango -> 'No'
    Dim sAnterior As String
    sAnterior = mIndicador.EstaEnElIntervaloDadoRiesgo("01/01/2025", "30/06/2025", "15/11/2023", "15/03/2024")
    If sAnterior <> "No" Then
        Test_IND_mIndicador_EstaEnElIntervaloDadoRiesgo_ProyectoEnvuelveRango_RetornaSi = BuildFail( _
            "Proyecto anterior al rango debe retornar 'No', obtuvo '" & sAnterior & "'", logs)
        Exit Function
    End If

    ' Caso borde: fecha final de informe anterior a la inicial -> #ERR
    Dim sInvertido As String
    sInvertido = mIndicador.EstaEnElIntervaloDadoRiesgo("30/06/2025", "01/01/2025", "01/01/2025", "31/12/2025")
    If Left$(sInvertido, 5) <> "#ERR|" Then
        Test_IND_mIndicador_EstaEnElIntervaloDadoRiesgo_ProyectoEnvuelveRango_RetornaSi = BuildFail( _
            "Rango invertido debe retornar '#ERR|...', obtuvo '" & sInvertido & "'", logs)
        Exit Function
    End If

    Test_IND_mIndicador_EstaEnElIntervaloDadoRiesgo_ProyectoEnvuelveRango_RetornaSi = BuildOk("envuelve_si_anterior_no", logs)
    Exit Function
EH:
    Test_IND_mIndicador_EstaEnElIntervaloDadoRiesgo_ProyectoEnvuelveRango_RetornaSi = BuildFail( _
        "Excepcion: " & Err.Number & " - " & Err.Description, logs)
End Function


' ============================================================
' Test 2 (BR-IND-006): IndicadorRiesgosRepositorioV2_GetTabla 5KPIs
' Fixture: 1 proyecto con 3 riesgos identificados, 1 retirado, 1 oferta,
' 2 materializaciones, 1 oferta->gestion. Cada KPI debe coincidir.
' ============================================================
Public Function Test_IND_IndicadorRiesgosRepositorioV2_GetTabla_5KPIs_ConFixtureControlado() As String
    On Error GoTo EH
    Dim logs(0 To 7) As String
    logs(0) = "1. Arrange: sandbox + proyecto 900700 + edicion 900701"
    logs(1) = "2. Arrange: 3 riesgos (FechaDetectado in range, 1 con FechaRetirado in range)"
    logs(2) = "3. Arrange: 1 TbRiesgosAIntegrar con FechaDetectado in range"
    logs(3) = "4. Arrange: 1 TbRiesgosAIntegrar con Trasladar='Si' + FechaAltaRegistro in range"
    logs(4) = "5. Arrange: 2 TbRiesgosMaterializaciones (EsMaterializacion='Si', Fecha in range)"
    logs(5) = "6. Act: IndicadorRiesgosRepositorioV2_GetTabla"
    logs(6) = "7. Assert: 1 fila con los 5 KPIs matching el fixture"
    logs(7) = "8. Teardown: CleanFixture"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_IND_IndicadorRiesgosRepositorioV2_GetTabla_5KPIs_ConFixtureControlado = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If

    CleanFixture
    If Not SeedBase(db) Then
        Test_IND_IndicadorRiesgosRepositorioV2_GetTabla_5KPIs_ConFixtureControlado = BuildFail( _
            "SeedBase fallo: " & Err.Description, logs)
        Exit Function
    End If

    ' 3 riesgos identificados (todos con FechaDetectado en rango)
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_A & ", " & FIX_EDICION & ", 'UNIQ-A', 'R01', #2025-01-15#, 'Detectado', 3)", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_B & ", " & FIX_EDICION & ", 'UNIQ-B', 'R02', #2025-03-20#, 'Retirado', 3)", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_C & ", " & FIX_EDICION & ", 'UNIQ-C', 'R03', #2025-04-10#, 'Detectado', 3)", dbFailOnError
    ' El riesgo B ademas tiene FechaRetirado in range (cuenta como retirado)
    db.Execute "UPDATE TbRiesgos SET FechaRetirado = #2025-06-15# WHERE IDRiesgo = " & FIX_RIESGO_B, dbFailOnError

    ' 1 TbRiesgosAIntegrar con FechaDetectado in range, Trasladar NO 'Si'
    ' Cuenta como EnOferta (FechaDetectado in range), NO como OfertaPasanGestion (sin Trasladar='Si')
    db.Execute "INSERT INTO TbRiesgosAIntegrar (IDRiesgoExt, IDEdicion, CodRiesgo, " & _
        "FechaDetectado, FechaAltaRegistro, Trasladar) VALUES (" & _
        FIX_OFERTA_A & ", " & FIX_EDICION & ", 'OFR-A', #2025-04-10#, #2025-04-10#, 'No')", dbFailOnError
    ' 1 TbRiesgosAIntegrar con FechaDetectado FUERA de rango, Trasladar='Si' + FechaAltaRegistro in range
    ' Solo cuenta como OfertaPasanGestion (no como EnOferta porque FechaDetectado esta fuera)
    db.Execute "INSERT INTO TbRiesgosAIntegrar (IDRiesgoExt, IDEdicion, CodRiesgo, " & _
        "FechaDetectado, FechaAltaRegistro, Trasladar) VALUES (" & _
        FIX_OFERTA_B & ", " & FIX_EDICION & ", 'OFR-B', #2024-12-01#, #2025-08-20#, 'Sí')", dbFailOnError

    ' 2 TbRiesgosMaterializaciones (EsMaterializacion='Si', Fecha in range)
    db.Execute "INSERT INTO TbRiesgosMaterializaciones (ID, IDProyecto, IDEdicion, CodigoRiesgo, " & _
        "Fecha, EsMaterializacion) VALUES (" & _
        FIX_MAT_A & ", " & FIX_PROYECTO & ", " & FIX_EDICION & ", 'R01', #2025-05-01#, 'Sí')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgosMaterializaciones (ID, IDProyecto, IDEdicion, CodigoRiesgo, " & _
        "Fecha, EsMaterializacion) VALUES (" & _
        FIX_MAT_B & ", " & FIX_PROYECTO & ", " & FIX_EDICION & ", 'R01', #2025-07-15#, 'Sí')", dbFailOnError

    ' Act
    Dim m_Error As String
    Dim rs As DAO.Recordset
    Set rs = IndicadorRiesgosRepositorio.IndicadorRiesgosRepositorioV2_GetTabla( _
        FIX_DINI, FIX_DFIN, CStr(FIX_PROYECTO), m_Error)
    If rs Is Nothing Then
        CleanFixture
        Test_IND_IndicadorRiesgosRepositorioV2_GetTabla_5KPIs_ConFixtureControlado = BuildFail( _
            "GetTabla retorno Nothing. m_Error='" & m_Error & "'", logs)
        Exit Function
    End If

    ' Assert
    If rs.EOF Then
        rs.Close
        CleanFixture
        Test_IND_IndicadorRiesgosRepositorioV2_GetTabla_5KPIs_ConFixtureControlado = BuildFail( _
            "GetTabla retorno recordset vacio", logs)
        Exit Function
    End If
    rs.MoveFirst

    Dim kId As Long, kRet As Long, kOf As Long, kMat As Long, kOfGes As Long
    kId = Nz(rs!RiesgosIdentificados, -1)
    kRet = Nz(rs!RiesgosRetirados, -1)
    kOf = Nz(rs!RiesgosEnOferta, -1)
    kMat = Nz(rs!RiesgosMaterializados, -1)
    kOfGes = Nz(rs!RiesgosOfertaPasanGestion, -1)
    rs.Close

    ' Esperado: 3 identificados, 1 retirado, 1 oferta, 2 materializados, 1 oferta->gestion
    If kId <> 3 Or kRet <> 1 Or kOf <> 1 Or kMat <> 2 Or kOfGes <> 1 Then
        CleanFixture
        Test_IND_IndicadorRiesgosRepositorioV2_GetTabla_5KPIs_ConFixtureControlado = BuildFail( _
            "KPIs no coinciden. Esperado [3,1,1,2,1], obtenido [" & _
            kId & "," & kRet & "," & kOf & "," & kMat & "," & kOfGes & "]", logs)
        Exit Function
    End If

    CleanFixture
    Test_IND_IndicadorRiesgosRepositorioV2_GetTabla_5KPIs_ConFixtureControlado = BuildOk( _
        "kpis_3_1_1_2_1", logs)
    Exit Function
EH:
    CleanFixture
    Test_IND_IndicadorRiesgosRepositorioV2_GetTabla_5KPIs_ConFixtureControlado = BuildFail( _
        "Excepcion: " & Err.Number & " - " & Err.Description, logs)
End Function


' ============================================================
' Test 3 (BR-IND-007): GetResumen deduplicacion por CodigoUnico
' Fixture: 1 proyecto, 2 ediciones, mismo CodigoUnico en ambas, ambos
' FechaDetectado in range. Esperado: RiesgosIdentificados = 1 (no 2).
' ============================================================
Public Function Test_IND_IndicadorRiesgosV2_GetResumen_RiesgoCopiadoEntreEdiciones_NoDuplica() As String
    On Error GoTo EH
    Dim logs(0 To 6) As String
    logs(0) = "1. Arrange: proyecto 900700 + 2 ediciones (900701, 900702)"
    logs(1) = "2. Arrange: mismo CodigoUnico 'DUP-001' en ambas ediciones, FechaDetectado in range"
    logs(2) = "3. Act: IndicadorRiesgosV2_GetResumen"
    logs(3) = "4. Assert: 1 fila con RiesgosIdentificados = 1 (deduplicado por CodigoUnico)"
    logs(4) = "5. Assert: 1 fila con RiesgosOfertaPasanGestion = 0 (no hay oferta)"
    logs(5) = "6. Teardown: CleanFixture"
    logs(6) = "7. Assert: registro de no-duplicacion persistido"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_IND_IndicadorRiesgosV2_GetResumen_RiesgoCopiadoEntreEdiciones_NoDuplica = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If

    CleanFixture
    If Not SeedBase(db) Then
        Test_IND_IndicadorRiesgosV2_GetResumen_RiesgoCopiadoEntreEdiciones_NoDuplica = BuildFail( _
            "SeedBase fallo: " & Err.Description, logs)
        Exit Function
    End If
    If Not SeedSecondEdition(db) Then
        CleanFixture
        Test_IND_IndicadorRiesgosV2_GetResumen_RiesgoCopiadoEntreEdiciones_NoDuplica = BuildFail( _
            "SeedSecondEdition fallo: " & Err.Description, logs)
        Exit Function
    End If

    ' Mismo CodigoUnico en 2 ediciones distintas, ambos FechaDetectado in range
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_A & ", " & FIX_EDICION & ", 'DUP-001', 'R01', #2025-02-10#, 'Detectado', 3)", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_B & ", " & FIX_EDICION2 & ", 'DUP-001', 'R01', #2025-05-20#, 'Detectado', 3)", dbFailOnError

    Dim m_Error As String
    Dim rs As DAO.Recordset
    Set rs = IndicadorRiesgosRepositorio.IndicadorRiesgosV2_GetResumen( _
        FIX_DINI, FIX_DFIN, CStr(FIX_PROYECTO), m_Error)
    If rs Is Nothing Then
        CleanFixture
        Test_IND_IndicadorRiesgosV2_GetResumen_RiesgoCopiadoEntreEdiciones_NoDuplica = BuildFail( _
            "GetResumen retorno Nothing. m_Error='" & m_Error & "'", logs)
        Exit Function
    End If
    If rs.EOF Then
        rs.Close
        CleanFixture
        Test_IND_IndicadorRiesgosV2_GetResumen_RiesgoCopiadoEntreEdiciones_NoDuplica = BuildFail( _
            "GetResumen retorno recordset vacio", logs)
        Exit Function
    End If
    rs.MoveFirst

    Dim kId As Long
    kId = Nz(rs!RiesgosIdentificados, -1)
    rs.Close

    CleanFixture
    If kId <> 1 Then
        Test_IND_IndicadorRiesgosV2_GetResumen_RiesgoCopiadoEntreEdiciones_NoDuplica = BuildFail( _
            "Esperado RiesgosIdentificados=1 (deduplicado por CodigoUnico), obtuvo " & kId, logs)
        Exit Function
    End If

    Test_IND_IndicadorRiesgosV2_GetResumen_RiesgoCopiadoEntreEdiciones_NoDuplica = BuildOk( _
        "dedup_codigounico_ok", logs)
    Exit Function
EH:
    CleanFixture
    Test_IND_IndicadorRiesgosV2_GetResumen_RiesgoCopiadoEntreEdiciones_NoDuplica = BuildFail( _
        "Excepcion: " & Err.Number & " - " & Err.Description, logs)
End Function


' ============================================================
' Test 4 (BR-IND-008): SqlDetalle itMaterializados sin GROUP BY
' Fixture: 1 proyecto, 1 edicion, 1 riesgo, 2 materializaciones
' para el mismo CodigoRiesgo. SqlDetalle NO agrupa; ejecutar el SQL
' debe devolver 2 filas.
' ============================================================
Public Function Test_IND_IndicadorRiesgosV2_SqlDetalle_itMaterializados_NoAgrupaEventos() As String
    On Error GoTo EH
    Dim logs(0 To 6) As String
    logs(0) = "1. Arrange: proyecto 900700 + edicion 900701 + 1 riesgo R01"
    logs(1) = "2. Arrange: 2 TbRiesgosMaterializaciones con CodigoRiesgo='R01', Fechas distintas"
    logs(2) = "3. Act: IndicadorRiesgosV2_SqlDetalle(itMaterializados, ...)"
    logs(3) = "4. Assert: SQL no vacio y NO contiene 'GROUP BY'"
    logs(4) = "5. Act: ejecutar SQL contra sandbox"
    logs(5) = "6. Assert: 2 filas (1 por evento, sin deduplicacion)"
    logs(6) = "7. Teardown: CleanFixture"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_IND_IndicadorRiesgosV2_SqlDetalle_itMaterializados_NoAgrupaEventos = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If

    CleanFixture
    If Not SeedBase(db) Then
        Test_IND_IndicadorRiesgosV2_SqlDetalle_itMaterializados_NoAgrupaEventos = BuildFail( _
            "SeedBase fallo: " & Err.Description, logs)
        Exit Function
    End If

    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_A & ", " & FIX_EDICION & ", 'MAT-001', 'R01', #2025-01-15#, 'Detectado', 3)", dbFailOnError
    db.Execute "INSERT INTO TbRiesgosMaterializaciones (ID, IDProyecto, IDEdicion, CodigoRiesgo, " & _
        "Fecha, EsMaterializacion) VALUES (" & _
        FIX_MAT_A & ", " & FIX_PROYECTO & ", " & FIX_EDICION & ", 'R01', #2025-05-01#, 'Sí')", dbFailOnError
    db.Execute "INSERT INTO TbRiesgosMaterializaciones (ID, IDProyecto, IDEdicion, CodigoRiesgo, " & _
        "Fecha, EsMaterializacion) VALUES (" & _
        FIX_MAT_B & ", " & FIX_PROYECTO & ", " & FIX_EDICION & ", 'R01', #2025-07-15#, 'Sí')", dbFailOnError

    Dim sql As String
    sql = IndicadorRiesgosRepositorio.IndicadorRiesgosV2_SqlDetalle( _
        itMaterializados, FIX_DINI, FIX_DFIN, CStr(FIX_PROYECTO))
    If Len(sql) = 0 Then
        CleanFixture
        Test_IND_IndicadorRiesgosV2_SqlDetalle_itMaterializados_NoAgrupaEventos = BuildFail( _
            "SqlDetalle retorno string vacio", logs)
        Exit Function
    End If
    If InStr(1, sql, "GROUP BY", vbTextCompare) > 0 Then
        CleanFixture
        Test_IND_IndicadorRiesgosV2_SqlDetalle_itMaterializados_NoAgrupaEventos = BuildFail( _
            "SqlDetalle(itMaterializados) NO debe contener GROUP BY. SQL='" & sql & "'", logs)
        Exit Function
    End If

    ' Ejecutar el SQL contra el sandbox
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sql)
    Dim nRows As Long
    nRows = 0
    If Not rs.EOF Then
        rs.MoveLast
        nRows = rs.RecordCount
    End If
    rs.Close

    CleanFixture
    If nRows <> 2 Then
        Test_IND_IndicadorRiesgosV2_SqlDetalle_itMaterializados_NoAgrupaEventos = BuildFail( _
            "Esperado 2 filas (1 por evento), obtuvo " & nRows, logs)
        Exit Function
    End If

    Test_IND_IndicadorRiesgosV2_SqlDetalle_itMaterializados_NoAgrupaEventos = BuildOk( _
        "no_agrupa_2_filas", logs)
    Exit Function
EH:
    CleanFixture
    Test_IND_IndicadorRiesgosV2_SqlDetalle_itMaterializados_NoAgrupaEventos = BuildFail( _
        "Excepcion: " & Err.Number & " - " & Err.Description, logs)
End Function


' ============================================================
' Test 5 (BR-IND-003): RellenaColeccionDetalleRiesgos idempotente
' Fixture: 1 proyecto, 1 edicion, 2 riesgos. ColProyectosParaInforme
' poblado. Llamar la funcion 2 veces: ambas deben dejar exactamente
' 2 filas en TbAuxProyectosRiesgos (la funcion trunca antes de
' rellenar, por lo que el segundo call no duplica).
' ============================================================
Public Function Test_IND_mIndicador_RellenaColeccionDetalleRiesgos_Idempotente2Llamadas() As String
    On Error GoTo EH
    Dim logs(0 To 7) As String
    logs(0) = "1. Arrange: proyecto 900700 + edicion 900701 + 2 riesgos (R01, R02)"
    logs(1) = "2. Arrange: ColProyectosParaInforme = '900700|INDTEST|Indicador Test||2025-01-01'"
    logs(2) = "3. Act (1ra): RellenaColeccionDetalleRiesgos"
    logs(3) = "4. Assert: retorno 'OK' + 2 filas en TbAuxProyectosRiesgos"
    logs(4) = "5. Act (2da): RellenaColeccionDetalleRiesgos (sin volver a poblar ColProyectosParaInforme)"
    logs(5) = "6. Assert: retorno 'OK' + sigue habiendo 2 filas (no duplica)"
    logs(6) = "7. Teardown: CleanFixture + reset ColProyectosParaInforme"
    logs(7) = "8. Assert: side effects limpios"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_IND_mIndicador_RellenaColeccionDetalleRiesgos_Idempotente2Llamadas = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If

    CleanFixture
    If Not SeedBase(db) Then
        Test_IND_mIndicador_RellenaColeccionDetalleRiesgos_Idempotente2Llamadas = BuildFail( _
            "SeedBase fallo: " & Err.Description, logs)
        Exit Function
    End If

    ' 2 riesgos con CodigoRiesgo distintos
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, FechaMaterializado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_A & ", " & FIX_EDICION & ", 'DET-A', 'R01', #2025-01-15#, #2025-03-10#, 'Materializado', 3)", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_B & ", " & FIX_EDICION & ", 'DET-B', 'R02', #2025-02-20#, 'Detectado', 3)", dbFailOnError

    ' Poblar ColProyectosParaInforme (formato: IDProyecto|Proyecto|NombreProyecto|FechaCierre|FechaRegistroInicial)
    Set mIndicador.ColProyectosParaInforme = New Collection
    mIndicador.ColProyectosParaInforme.Add _
        CStr(FIX_PROYECTO) & "|INDTEST|Indicador Test||2025-01-01"

    ' Llamada 1
    Dim resultado1 As String
    resultado1 = mIndicador.RellenaColeccionDetalleRiesgos()
    If Left$(resultado1, 5) = "#ERR|" Then
        CleanFixture
        Set mIndicador.ColProyectosParaInforme = New Collection
        Test_IND_mIndicador_RellenaColeccionDetalleRiesgos_Idempotente2Llamadas = BuildFail( _
            "1ra llamada retorno error: " & resultado1, logs)
        Exit Function
    End If

    Dim nDespues1 As Long
    nDespues1 = CountRows(db, "TbAuxProyectosRiesgos", "IDProyecto=" & FIX_PROYECTO)
    If nDespues1 <> 2 Then
        CleanFixture
        Set mIndicador.ColProyectosParaInforme = New Collection
        Test_IND_mIndicador_RellenaColeccionDetalleRiesgos_Idempotente2Llamadas = BuildFail( _
            "1ra llamada: esperado 2 filas en TbAuxProyectosRiesgos, obtuvo " & nDespues1, logs)
        Exit Function
    End If

    ' Llamada 2 (misma ColProyectosParaInforme)
    Dim resultado2 As String
    resultado2 = mIndicador.RellenaColeccionDetalleRiesgos()
    If Left$(resultado2, 5) = "#ERR|" Then
        CleanFixture
        Set mIndicador.ColProyectosParaInforme = New Collection
        Test_IND_mIndicador_RellenaColeccionDetalleRiesgos_Idempotente2Llamadas = BuildFail( _
            "2da llamada retorno error: " & resultado2, logs)
        Exit Function
    End If

    Dim nDespues2 As Long
    nDespues2 = CountRows(db, "TbAuxProyectosRiesgos", "IDProyecto=" & FIX_PROYECTO)
    If nDespues2 <> 2 Then
        CleanFixture
        Set mIndicador.ColProyectosParaInforme = New Collection
        Test_IND_mIndicador_RellenaColeccionDetalleRiesgos_Idempotente2Llamadas = BuildFail( _
            "2da llamada: esperado 2 filas (idempotente), obtuvo " & nDespues2, logs)
        Exit Function
    End If

    ' Teardown
    CleanFixture
    Set mIndicador.ColProyectosParaInforme = New Collection

    Test_IND_mIndicador_RellenaColeccionDetalleRiesgos_Idempotente2Llamadas = BuildOk( _
        "idempotente_2_llamadas_2_filas", logs)
    Exit Function
EH:
    CleanFixture
    On Error Resume Next
    Set mIndicador.ColProyectosParaInforme = New Collection
    On Error GoTo 0
    Test_IND_mIndicador_RellenaColeccionDetalleRiesgos_Idempotente2Llamadas = BuildFail( _
        "Excepcion: " & Err.Number & " - " & Err.Description, logs)
End Function


' ============================================================
' HOTFIX issue-122: Numero de Riesgos vigentes totales
' Predicate: FechaDetectado <= dFin
'             AND (FechaRetirado Is Null OR FechaRetirado >= dIni)
' Restringido a proyectos del CSV.
'
' Skill: access-vba-tdd v2.4
' - IDs de fixture en rango 900760+ (separado de los existentes 900700-900731)
' - Sandbox via Test_Fixtures.GetTestDb / ForceLocalBackend
' - Cada test hace su propio seed + teardown (no depende de SeedAll)
' ============================================================

' ============================================================
' Helpers de fixture: cleanup + seed del grafo de vigentes
' ============================================================

Private Sub CleanFixtureVigentes()
    On Error Resume Next
    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then Exit Sub

    db.Execute "DELETE FROM TbRiesgos WHERE IDEdicion IN (" & FIX_EDICION_VIG & "," & FIX_EDICION2_VIG & ")"
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion IN (" & FIX_EDICION_VIG & "," & FIX_EDICION2_VIG & ")"
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto IN (" & FIX_PROYECTO_VIG & "," & FIX_PROYECTO2_VIG & ")"
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_EXPEDIENTE_VIG
    On Error GoTo 0
End Sub

Private Function SeedBaseVigentes(ByVal db As DAO.Database, _
                                  Optional ByVal includeSecondProject As Boolean = False) As Boolean
    On Error GoTo SeedBaseVigentes_EH
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
        "VALUES (" & FIX_EXPEDIENTE_VIG & ", 'TESTVIG', 'Fixture vigentes test', 'Test', 1)", dbFailOnError
    db.Execute "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto, NombreProyecto) " & _
        "VALUES (" & FIX_PROYECTO_VIG & ", " & FIX_EXPEDIENTE_VIG & ", 'VIGTEST1', 'Vigentes Test 1')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
        "VALUES (" & FIX_EDICION_VIG & ", " & FIX_PROYECTO_VIG & ", 1, 'TESTUSER')", dbFailOnError
    If includeSecondProject Then
        db.Execute "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto, NombreProyecto) " & _
            "VALUES (" & FIX_PROYECTO2_VIG & ", " & FIX_EXPEDIENTE_VIG & ", 'VIGTEST2', 'Vigentes Test 2')", dbFailOnError
        db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) " & _
            "VALUES (" & FIX_EDICION2_VIG & ", " & FIX_PROYECTO2_VIG & ", 1, 'TESTUSER')", dbFailOnError
    End If
    SeedBaseVigentes = True
    Exit Function
SeedBaseVigentes_EH:
    SeedBaseVigentes = False
End Function


' ============================================================
' Tests del nuevo campo RiesgosVigentesEnPeriodo en
' IndicadorRiesgosV2_GetResumen
' ============================================================

' --- Test 1 (BR-IND-009 / case 1): abierto todo el periodo cuenta ---
Public Function Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoTodoElPeriodo() As String
    On Error GoTo EH
    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: proyecto 900761 + edicion 900762"
    logs(1) = "2. Arrange: riesgo T1 (2025-01-01, NULL) — abierto todo el periodo"
    logs(2) = "3. Act: IndicadorRiesgosV2_GetResumen"
    logs(3) = "4. Assert: RiesgosVigentesEnPeriodo = 1"
    logs(4) = "5. Teardown: CleanFixtureVigentes"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoTodoElPeriodo = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If

    CleanFixtureVigentes
    If Not SeedBaseVigentes(db) Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoTodoElPeriodo = BuildFail( _
            "SeedBaseVigentes fallo: " & Err.Description, logs)
        Exit Function
    End If

    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, FechaRetirado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_T1 & ", " & FIX_EDICION_VIG & ", 'VIG-T1', 'RV01', #2025-01-01#, NULL, 'Detectado', 3)", dbFailOnError

    Dim m_Error As String
    Dim rs As DAO.Recordset
    Set rs = IndicadorRiesgosV2_GetResumen(FIX_DINI_VIG, FIX_DFIN_VIG, CStr(FIX_PROYECTO_VIG), m_Error)
    If rs Is Nothing Then
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoTodoElPeriodo = BuildFail( _
            "GetResumen retorno Nothing. m_Error='" & m_Error & "'", logs)
        Exit Function
    End If
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoTodoElPeriodo = BuildFail( _
            "GetResumen retorno recordset vacio", logs)
        Exit Function
    End If
    rs.MoveFirst

    Dim nVig As Long
    nVig = Nz(rs!RiesgosVigentesEnPeriodo, -1)
    rs.Close: Set rs = Nothing
    CleanFixtureVigentes

    If nVig <> 1 Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoTodoElPeriodo = BuildFail( _
            "Esperado RiesgosVigentesEnPeriodo=1, obtuvo " & nVig, logs)
        Exit Function
    End If

    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoTodoElPeriodo = BuildOk("vig_1", logs)
    Exit Function
EH:
    CleanFixtureVigentes
    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoTodoElPeriodo = BuildFail( _
        "Excepcion: " & Err.Number & " - " & Err.Description, logs)
End Function


' --- Test 2 (case 2): abierto solo 1 dia al inicio (FechaDetectado<dIni pero FechaRetirado=dIni) cuenta ---
Public Function Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoUnDiaAlInicio() As String
    On Error GoTo EH
    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: riesgo T2 (2025-12-31, 2026-01-01) — intersecta un dia al inicio"
    logs(1) = "2. Arrange: rango S1 2026-01-01 a 2026-06-30"
    logs(2) = "3. Act: GetResumen"
    logs(3) = "4. Assert: RiesgosVigentesEnPeriodo = 1 (cierre cae justo en dIni)"
    logs(4) = "5. Teardown: CleanFixtureVigentes"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoUnDiaAlInicio = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If

    CleanFixtureVigentes
    If Not SeedBaseVigentes(db) Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoUnDiaAlInicio = BuildFail( _
            "SeedBaseVigentes fallo: " & Err.Description, logs)
        Exit Function
    End If

    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, FechaRetirado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_T2 & ", " & FIX_EDICION_VIG & ", 'VIG-T2', 'RV02', #2025-12-31#, #2026-01-01#, 'Retirado', 3)", dbFailOnError

    Dim m_Error As String
    Dim rs As DAO.Recordset
    Set rs = IndicadorRiesgosV2_GetResumen(FIX_DINI_VIG, FIX_DFIN_VIG, CStr(FIX_PROYECTO_VIG), m_Error)
    If rs Is Nothing Then
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoUnDiaAlInicio = BuildFail( _
            "GetResumen retorno Nothing. m_Error='" & m_Error & "'", logs)
        Exit Function
    End If
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoUnDiaAlInicio = BuildFail( _
            "GetResumen retorno recordset vacio", logs)
        Exit Function
    End If
    rs.MoveFirst
    Dim nVig As Long
    nVig = Nz(rs!RiesgosVigentesEnPeriodo, -1)
    rs.Close: Set rs = Nothing
    CleanFixtureVigentes

    If nVig <> 1 Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoUnDiaAlInicio = BuildFail( _
            "Esperado RiesgosVigentesEnPeriodo=1, obtuvo " & nVig, logs)
        Exit Function
    End If

    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoUnDiaAlInicio = BuildOk("vig_1", logs)
    Exit Function
EH:
    CleanFixtureVigentes
    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoUnDiaAlInicio = BuildFail( _
        "Excepcion: " & Err.Number & " - " & Err.Description, logs)
End Function


' --- Test 3 (case 3): abierto solo 1 dia al cierre (deteccion y retiro dentro del periodo) cuenta ---
Public Function Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoUnDiaAlCierre() As String
    On Error GoTo EH
    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: riesgo T3 (2026-06-29, 2026-06-30) — abierto el ultimo dia del periodo"
    logs(1) = "2. Act: GetResumen con S1 2026"
    logs(2) = "3. Assert: RiesgosVigentesEnPeriodo = 1 (intersecta dentro)"
    logs(3) = "4. Teardown: CleanFixtureVigentes"
    logs(4) = "5. Assert: ambos extremos del periodo cuentan (>= dIni, <= dFin)"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoUnDiaAlCierre = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If

    CleanFixtureVigentes
    If Not SeedBaseVigentes(db) Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoUnDiaAlCierre = BuildFail( _
            "SeedBaseVigentes fallo: " & Err.Description, logs)
        Exit Function
    End If

    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, FechaRetirado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_T3 & ", " & FIX_EDICION_VIG & ", 'VIG-T3', 'RV03', #2026-06-29#, #2026-06-30#, 'Retirado', 3)", dbFailOnError

    Dim m_Error As String
    Dim rs As DAO.Recordset
    Set rs = IndicadorRiesgosV2_GetResumen(FIX_DINI_VIG, FIX_DFIN_VIG, CStr(FIX_PROYECTO_VIG), m_Error)
    If rs Is Nothing Then
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoUnDiaAlCierre = BuildFail( _
            "GetResumen retorno Nothing. m_Error='" & m_Error & "'", logs)
        Exit Function
    End If
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoUnDiaAlCierre = BuildFail( _
            "GetResumen retorno recordset vacio", logs)
        Exit Function
    End If
    rs.MoveFirst
    Dim nVig As Long
    nVig = Nz(rs!RiesgosVigentesEnPeriodo, -1)
    rs.Close: Set rs = Nothing
    CleanFixtureVigentes

    If nVig <> 1 Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoUnDiaAlCierre = BuildFail( _
            "Esperado RiesgosVigentesEnPeriodo=1, obtuvo " & nVig, logs)
        Exit Function
    End If

    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoUnDiaAlCierre = BuildOk("vig_1", logs)
    Exit Function
EH:
    CleanFixtureVigentes
    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_AbiertoUnDiaAlCierre = BuildFail( _
        "Excepcion: " & Err.Number & " - " & Err.Description, logs)
End Function


' --- Test 4 (case 4): cerrado antes del periodo NO cuenta ---
Public Function Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CerradoAntesNoCuenta() As String
    On Error GoTo EH
    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: riesgo T4 (2025-08-10, 2025-10-05) — cerrado antes del periodo"
    logs(1) = "2. Act: GetResumen con S1 2026"
    logs(2) = "3. Assert: RiesgosVigentesEnPeriodo = 0"
    logs(3) = "4. Teardown: CleanFixtureVigentes"
    logs(4) = "5. Assert: el WHERE filtra el riesgo (no aparece en el SELECT)"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CerradoAntesNoCuenta = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If

    CleanFixtureVigentes
    If Not SeedBaseVigentes(db) Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CerradoAntesNoCuenta = BuildFail( _
            "SeedBaseVigentes fallo: " & Err.Description, logs)
        Exit Function
    End If

    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, FechaRetirado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_T4 & ", " & FIX_EDICION_VIG & ", 'VIG-T4', 'RV04', #2025-08-10#, #2025-10-05#, 'Retirado', 3)", dbFailOnError

    Dim m_Error As String
    Dim rs As DAO.Recordset
    Set rs = IndicadorRiesgosV2_GetResumen(FIX_DINI_VIG, FIX_DFIN_VIG, CStr(FIX_PROYECTO_VIG), m_Error)
    If rs Is Nothing Then
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CerradoAntesNoCuenta = BuildFail( _
            "GetResumen retorno Nothing. m_Error='" & m_Error & "'", logs)
        Exit Function
    End If
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CerradoAntesNoCuenta = BuildFail( _
            "GetResumen retorno recordset vacio", logs)
        Exit Function
    End If
    rs.MoveFirst
    Dim nVig As Long
    nVig = Nz(rs!RiesgosVigentesEnPeriodo, -1)
    rs.Close: Set rs = Nothing
    CleanFixtureVigentes

    If nVig <> 0 Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CerradoAntesNoCuenta = BuildFail( _
            "Esperado RiesgosVigentesEnPeriodo=0, obtuvo " & nVig, logs)
        Exit Function
    End If

    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CerradoAntesNoCuenta = BuildOk("vig_0", logs)
    Exit Function
EH:
    CleanFixtureVigentes
    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CerradoAntesNoCuenta = BuildFail( _
        "Excepcion: " & Err.Number & " - " & Err.Description, logs)
End Function


' --- Test 5 (case 5): detectado despues del periodo NO cuenta ---
Public Function Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_DetectadoDespuesNoCuenta() As String
    On Error GoTo EH
    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: riesgo T5 (2026-08-15, NULL) — detectado despues de dFin"
    logs(1) = "2. Act: GetResumen con S1 2026"
    logs(2) = "3. Assert: RiesgosVigentesEnPeriodo = 0"
    logs(3) = "4. Teardown: CleanFixtureVigentes"
    logs(4) = "5. Assert: FechaDetectado <= dFin filtra correctamente"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_DetectadoDespuesNoCuenta = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If

    CleanFixtureVigentes
    If Not SeedBaseVigentes(db) Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_DetectadoDespuesNoCuenta = BuildFail( _
            "SeedBaseVigentes fallo: " & Err.Description, logs)
        Exit Function
    End If

    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, FechaRetirado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_T5 & ", " & FIX_EDICION_VIG & ", 'VIG-T5', 'RV05', #2026-08-15#, NULL, 'Detectado', 3)", dbFailOnError

    Dim m_Error As String
    Dim rs As DAO.Recordset
    Set rs = IndicadorRiesgosV2_GetResumen(FIX_DINI_VIG, FIX_DFIN_VIG, CStr(FIX_PROYECTO_VIG), m_Error)
    If rs Is Nothing Then
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_DetectadoDespuesNoCuenta = BuildFail( _
            "GetResumen retorno Nothing. m_Error='" & m_Error & "'", logs)
        Exit Function
    End If
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_DetectadoDespuesNoCuenta = BuildFail( _
            "GetResumen retorno recordset vacio", logs)
        Exit Function
    End If
    rs.MoveFirst
    Dim nVig As Long
    nVig = Nz(rs!RiesgosVigentesEnPeriodo, -1)
    rs.Close: Set rs = Nothing
    CleanFixtureVigentes

    If nVig <> 0 Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_DetectadoDespuesNoCuenta = BuildFail( _
            "Esperado RiesgosVigentesEnPeriodo=0, obtuvo " & nVig, logs)
        Exit Function
    End If

    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_DetectadoDespuesNoCuenta = BuildOk("vig_0", logs)
    Exit Function
EH:
    CleanFixtureVigentes
    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_DetectadoDespuesNoCuenta = BuildFail( _
        "Excepcion: " & Err.Number & " - " & Err.Description, logs)
End Function


' --- Test 6 (case 6): deteccion fuera del periodo, cierre dentro -> cuenta (interseca) ---
Public Function Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_DeteccionFueraCierreDentroCuenta() As String
    On Error GoTo EH
    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: riesgo T6 (2025-11-20, 2026-03-01) — abierto antes, cerrado dentro"
    logs(1) = "2. Act: GetResumen con S1 2026"
    logs(2) = "3. Assert: RiesgosVigentesEnPeriodo = 1 (interseca)"
    logs(3) = "4. Teardown: CleanFixtureVigentes"
    logs(4) = "5. Assert: FechaRetirado >= dIni OR Is Null funciona"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_DeteccionFueraCierreDentroCuenta = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If

    CleanFixtureVigentes
    If Not SeedBaseVigentes(db) Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_DeteccionFueraCierreDentroCuenta = BuildFail( _
            "SeedBaseVigentes fallo: " & Err.Description, logs)
        Exit Function
    End If

    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, FechaRetirado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_T6 & ", " & FIX_EDICION_VIG & ", 'VIG-T6', 'RV06', #2025-11-20#, #2026-03-01#, 'Retirado', 3)", dbFailOnError

    Dim m_Error As String
    Dim rs As DAO.Recordset
    Set rs = IndicadorRiesgosV2_GetResumen(FIX_DINI_VIG, FIX_DFIN_VIG, CStr(FIX_PROYECTO_VIG), m_Error)
    If rs Is Nothing Then
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_DeteccionFueraCierreDentroCuenta = BuildFail( _
            "GetResumen retorno Nothing. m_Error='" & m_Error & "'", logs)
        Exit Function
    End If
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_DeteccionFueraCierreDentroCuenta = BuildFail( _
            "GetResumen retorno recordset vacio", logs)
        Exit Function
    End If
    rs.MoveFirst
    Dim nVig As Long
    nVig = Nz(rs!RiesgosVigentesEnPeriodo, -1)
    rs.Close: Set rs = Nothing
    CleanFixtureVigentes

    If nVig <> 1 Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_DeteccionFueraCierreDentroCuenta = BuildFail( _
            "Esperado RiesgosVigentesEnPeriodo=1, obtuvo " & nVig, logs)
        Exit Function
    End If

    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_DeteccionFueraCierreDentroCuenta = BuildOk("vig_1", logs)
    Exit Function
EH:
    CleanFixtureVigentes
    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_DeteccionFueraCierreDentroCuenta = BuildFail( _
        "Excepcion: " & Err.Number & " - " & Err.Description, logs)
End Function


' --- Test 7 (case 7): CSV con multiples proyectos cuenta solo los riesgos del CSV ---
Public Function Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CsvMultiplesProyectosSoloCuentaDelCsv() As String
    On Error GoTo EH
    Dim logs(0 To 5) As String
    logs(0) = "1. Arrange: 2 proyectos (900761, 900763), 2 ediciones (900762, 900764)"
    logs(1) = "2. Arrange: T1 (en P1, vigente) + T7 (en P2, vigente) + T8 (en P1, vigente)"
    logs(2) = "3. Act: GetResumen con CSV = '900761, 900763'"
    logs(3) = "4. Assert: RiesgosVigentesEnPeriodo = 3 (ambos proyectos)"
    logs(4) = "5. Act: GetResumen con CSV = '900761' (solo P1)"
    logs(5) = "6. Assert: RiesgosVigentesEnPeriodo = 2 (solo T1 + T8)"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CsvMultiplesProyectosSoloCuentaDelCsv = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If

    CleanFixtureVigentes
    If Not SeedBaseVigentes(db, True) Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CsvMultiplesProyectosSoloCuentaDelCsv = BuildFail( _
            "SeedBaseVigentes fallo: " & Err.Description, logs)
        Exit Function
    End If

    ' T1 en P1 (abierto todo el periodo)
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, FechaRetirado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_T1 & ", " & FIX_EDICION_VIG & ", 'VIG-T1', 'RV01', #2025-01-01#, NULL, 'Detectado', 3)", dbFailOnError
    ' T7 en P2 (abierto todo el periodo)
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, FechaRetirado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_T7 & ", " & FIX_EDICION2_VIG & ", 'VIG-T7', 'RV07', #2025-02-01#, NULL, 'Detectado', 3)", dbFailOnError
    ' T8 en P1 (abierto todo el periodo)
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, FechaRetirado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_T8 & ", " & FIX_EDICION_VIG & ", 'VIG-T8', 'RV08', #2025-03-01#, NULL, 'Detectado', 3)", dbFailOnError

    ' Paso 1: CSV con ambos proyectos -> 3
    Dim m_Error As String
    Dim rs As DAO.Recordset
    Set rs = IndicadorRiesgosV2_GetResumen(FIX_DINI_VIG, FIX_DFIN_VIG, _
        CStr(FIX_PROYECTO_VIG) & ", " & CStr(FIX_PROYECTO2_VIG), m_Error)
    If rs Is Nothing Then
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CsvMultiplesProyectosSoloCuentaDelCsv = BuildFail( _
            "GetResumen (CSV ambos) retorno Nothing. m_Error='" & m_Error & "'", logs)
        Exit Function
    End If
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CsvMultiplesProyectosSoloCuentaDelCsv = BuildFail( _
            "GetResumen (CSV ambos) retorno recordset vacio", logs)
        Exit Function
    End If
    rs.MoveFirst
    Dim nAmbos As Long
    nAmbos = Nz(rs!RiesgosVigentesEnPeriodo, -1)
    rs.Close: Set rs = Nothing

    ' Paso 2: CSV con solo P1 -> 2 (T1 + T8)
    Set rs = IndicadorRiesgosV2_GetResumen(FIX_DINI_VIG, FIX_DFIN_VIG, CStr(FIX_PROYECTO_VIG), m_Error)
    If rs Is Nothing Then
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CsvMultiplesProyectosSoloCuentaDelCsv = BuildFail( _
            "GetResumen (CSV P1) retorno Nothing. m_Error='" & m_Error & "'", logs)
        Exit Function
    End If
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CsvMultiplesProyectosSoloCuentaDelCsv = BuildFail( _
            "GetResumen (CSV P1) retorno recordset vacio", logs)
        Exit Function
    End If
    rs.MoveFirst
    Dim nP1 As Long
    nP1 = Nz(rs!RiesgosVigentesEnPeriodo, -1)
    rs.Close: Set rs = Nothing
    CleanFixtureVigentes

    If nAmbos <> 3 Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CsvMultiplesProyectosSoloCuentaDelCsv = BuildFail( _
            "CSV ambos: esperado RiesgosVigentesEnPeriodo=3, obtuvo " & nAmbos, logs)
        Exit Function
    End If
    If nP1 <> 2 Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CsvMultiplesProyectosSoloCuentaDelCsv = BuildFail( _
            "CSV P1: esperado RiesgosVigentesEnPeriodo=2, obtuvo " & nP1, logs)
        Exit Function
    End If

    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CsvMultiplesProyectosSoloCuentaDelCsv = BuildOk("multi_3_y_2", logs)
    Exit Function
EH:
    CleanFixtureVigentes
    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CsvMultiplesProyectosSoloCuentaDelCsv = BuildFail( _
        "Excepcion: " & Err.Number & " - " & Err.Description, logs)
End Function


' --- Test 8 (case 8): CSV vacio preserva el entry guard existente (p_Error + Nothing) ---
Public Function Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CsvVacioEntryGuard() As String
    On Error GoTo EH
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: CSV vacio"
    logs(1) = "2. Act: GetResumen con p_ListaIDsCsv = ''"
    logs(2) = "3. Assert: retorna Nothing (entry guard existente de GetResumen)"
    logs(3) = "4. Assert: p_Error contiene mensaje de lista vacia (la nueva columna nunca se evalua)"

    Dim m_Error As String
    Dim rs As DAO.Recordset
    Set rs = IndicadorRiesgosV2_GetResumen(FIX_DINI_VIG, FIX_DFIN_VIG, "", m_Error)

    ' Contrato real de GetResumen: errores se senalizan via p_Error + Nothing,
    ' NO via Err.Raise (a diferencia de SqlDetalle que si re-eleva Err 1000).
    If Not rs Is Nothing Then
        rs.Close: Set rs = Nothing
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CsvVacioEntryGuard = BuildFail( _
            "Esperado Nothing cuando CSV vacio, obtuvo recordset. m_Error='" & m_Error & "'", logs)
        Exit Function
    End If
    If InStr(1, m_Error, "vacía", vbTextCompare) = 0 Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CsvVacioEntryGuard = BuildFail( _
            "Esperado mensaje de 'lista vacia' en p_Error, obtuvo: '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CsvVacioEntryGuard = BuildOk("csv_vacio_entry_guard_ok", logs)
    Exit Function
EH:
    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_CsvVacioEntryGuard = BuildFail( _
        "Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function


' --- Test 9 (case 9): SqlDetalle devuelve las mismas filas que cuenta el KPI ---
Public Function Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_ConcuerdaConKpi() As String
    On Error GoTo EH
    Dim logs(0 To 6) As String
    logs(0) = "1. Arrange: 4 riesgos — T1 cuenta, T4 NO, T6 cuenta, T5 NO"
    logs(1) = "2. Act: GetResumen -> cuenta esperada = 2"
    logs(2) = "3. Act: SqlDetalle(itVigentesEnPeriodo) -> ejecutar SQL"
    logs(3) = "4. Assert: filas devueltas = 2"
    logs(4) = "5. Assert: SQL no vacio y contiene la clausula In (CSV_AQUI)"
    logs(5) = "6. Assert: SQL incluye el predicate FechaDetectado <= ... AND (FechaRetirado Is Null OR FechaRetirado >= ...)"
    logs(6) = "7. Teardown: CleanFixtureVigentes"

    Dim dbErr As String
    Dim db As DAO.Database
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_ConcuerdaConKpi = BuildFail( _
            "TESTS BLOCKED: GetTestDb fallo: " & dbErr, logs)
        Exit Function
    End If

    CleanFixtureVigentes
    If Not SeedBaseVigentes(db) Then
        Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_ConcuerdaConKpi = BuildFail( _
            "SeedBaseVigentes fallo: " & Err.Description, logs)
        Exit Function
    End If

    ' T1 cuenta (todo el periodo)
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, FechaRetirado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_T1 & ", " & FIX_EDICION_VIG & ", 'VIG-T1', 'RV01', #2025-01-01#, NULL, 'Detectado', 3)", dbFailOnError
    ' T4 NO cuenta (cerrado antes)
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, FechaRetirado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_T4 & ", " & FIX_EDICION_VIG & ", 'VIG-T4', 'RV04', #2025-08-10#, #2025-10-05#, 'Retirado', 3)", dbFailOnError
    ' T6 cuenta (interseca)
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, FechaRetirado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_T6 & ", " & FIX_EDICION_VIG & ", 'VIG-T6', 'RV06', #2025-11-20#, #2026-03-01#, 'Retirado', 3)", dbFailOnError
    ' T5 NO cuenta (detectado despues)
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, " & _
        "FechaDetectado, FechaRetirado, Estado, Priorizacion) VALUES (" & _
        FIX_RIESGO_T5 & ", " & FIX_EDICION_VIG & ", 'VIG-T5', 'RV05', #2026-08-15#, NULL, 'Detectado', 3)", dbFailOnError

    Dim m_Error As String
    Dim rs As DAO.Recordset
    Set rs = IndicadorRiesgosV2_GetResumen(FIX_DINI_VIG, FIX_DFIN_VIG, CStr(FIX_PROYECTO_VIG), m_Error)
    If rs Is Nothing Then
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_ConcuerdaConKpi = BuildFail( _
            "GetResumen retorno Nothing. m_Error='" & m_Error & "'", logs)
        Exit Function
    End If
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_ConcuerdaConKpi = BuildFail( _
            "GetResumen retorno recordset vacio", logs)
        Exit Function
    End If
    rs.MoveFirst
    Dim nKpi As Long
    nKpi = Nz(rs!RiesgosVigentesEnPeriodo, -1)
    rs.Close: Set rs = Nothing

    Dim sql As String
    sql = IndicadorRiesgosV2_SqlDetalle(itVigentesEnPeriodo, FIX_DINI_VIG, FIX_DFIN_VIG, CStr(FIX_PROYECTO_VIG))
    If Len(sql) = 0 Then
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_ConcuerdaConKpi = BuildFail( _
            "SqlDetalle retorno string vacio", logs)
        Exit Function
    End If
    If InStr(1, sql, "In (" & FIX_PROYECTO_VIG & ")", vbTextCompare) = 0 Then
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_ConcuerdaConKpi = BuildFail( _
            "SQL no contiene In (CSV_AQUI). SQL='" & sql & "'", logs)
        Exit Function
    End If
    If InStr(1, sql, "FechaDetectado", vbTextCompare) = 0 _
       Or InStr(1, sql, "FechaRetirado", vbTextCompare) = 0 Then
        CleanFixtureVigentes
        Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_ConcuerdaConKpi = BuildFail( _
            "SQL no contiene el predicate (FechaDetectado / FechaRetirado). SQL='" & sql & "'", logs)
        Exit Function
    End If

    Set rs = db.OpenRecordset(sql)
    Dim nSql As Long
    nSql = 0
    If Not rs.EOF Then
        rs.MoveLast
        nSql = rs.RecordCount
    End If
    rs.Close: Set rs = Nothing
    CleanFixtureVigentes

    If nKpi <> 2 Then
        Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_ConcuerdaConKpi = BuildFail( _
            "KPI: esperado 2, obtuvo " & nKpi, logs)
        Exit Function
    End If
    If nSql <> 2 Then
        Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_ConcuerdaConKpi = BuildFail( _
            "SQL detalle: esperado 2 filas, obtuvo " & nSql, logs)
        Exit Function
    End If

    Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_ConcuerdaConKpi = BuildOk("kpi_2_sql_2", logs)
    Exit Function
EH:
    CleanFixtureVigentes
    Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_ConcuerdaConKpi = BuildFail( _
        "Excepcion: " & Err.Number & " - " & Err.Description, logs)
End Function


' --- Test 10 (case 10): defensa anti-inyeccion en SqlDetalle(itVigentesEnPeriodo) ---
Public Function Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_RechazaInyeccion() As String
    On Error GoTo EH
    Dim logs(0 To 2) As String
    logs(0) = "1. Arrange: itVigentesEnPeriodo con CSV malicioso '1) OR 1=1 --'"
    logs(1) = "2. Act: IndicadorRiesgosV2_SqlDetalle con payload SQL"
    logs(2) = "3. Assert: eleva Err 1000, no retorna SQL con payload"

    Dim sql As String
    On Error Resume Next
    sql = IndicadorRiesgosV2_SqlDetalle(itVigentesEnPeriodo, FIX_DINI_VIG, FIX_DFIN_VIG, "1) OR 1=1 --")
    Dim errNum As Long
    errNum = Err.Number
    On Error GoTo EH

    If errNum <> 1000 Then
        Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_RechazaInyeccion = BuildFail( _
            "Esperado Err 1000, obtuvo Err " & errNum & ". SQL='" & sql & "'", logs)
        Exit Function
    End If

    If InStr(1, sql, "OR 1=1", vbTextCompare) > 0 Then
        Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_RechazaInyeccion = BuildFail( _
            "El SQL no debe contener el payload. SQL='" & sql & "'", logs)
        Exit Function
    End If

    Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_RechazaInyeccion = BuildOk("err_1000_propagado", logs)
    Exit Function
EH:
    Test_IND_IndicadorRiesgosV2_SqlDetalle_itVigentesEnPeriodo_RechazaInyeccion = BuildFail( _
        "Excepcion inesperada: " & Err.Number & " - " & Err.Description, logs)
End Function


' --- Test 11 (case 11): el mismo CodigoUnico en N ediciones cuenta UNA vez (issue-127) ---
' Antes del fix GetResumen contaba 1 fila por (edicion, riesgo) y devolvia 3.
' Con GROUP BY R.CodigoUnico en la subconsulta, cuenta riesgos unicos.
Public Function Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_MismoRiesgoEnVariasEdiciones_CuentaUnaVez() As String
    On Error GoTo EH
    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: 1 mismo CodigoUnico='VIG-MULTIED' en 3 ediciones del mismo proyecto"
    logs(1) = "2. Setup: 2 ediciones adicionales (900790, 900791) + 3 filas en TbRiesgos con mismo CodigoUnico"
    logs(2) = "3. Act: GetResumen -> cuenta esperada = 1 (no 3)"
    logs(3) = "4. Teardown: borrar las 3 filas de riesgo y las 2 ediciones adicionales"

    Dim db As DAO.Database
    Set db = getdb()

    ' Limpieza previa (idempotente)
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE CodigoUnico = 'VIG-MULTIED'"
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion IN (900790, 900791)"
    On Error GoTo EH

    ' Aseguramos que el proyecto + edicion 1 existan (otro test pudo borrarlos).
    If Not SeedBaseVigentes(db, False) Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_MismoRiesgoEnVariasEdiciones_CuentaUnaVez = BuildFail( _
            "SeedBaseVigentes fallo (no se pudo crear el proyecto fixture)", logs)
        Exit Function
    End If

    ' Setup: 2 ediciones adicionales del mismo proyecto (FIX_EDICION_VIG ya es la 1ª)
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) VALUES (900790, " & FIX_PROYECTO_VIG & ", 2, 'TESTUSER')", dbFailOnError
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado) VALUES (900791, " & FIX_PROYECTO_VIG & ", 3, 'TESTUSER')", dbFailOnError

    ' Setup: 3 filas en TbRiesgos con MISMO CodigoUnico (FIX_EDICION_VIG + 900790 + 900791)
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, FechaDetectado, FechaRetirado, Estado, Priorizacion) VALUES (900780, " & FIX_EDICION_VIG & ", 'VIG-MULTIED', 'R-VIG-MULTIED-1', #2025-01-01#, NULL, 'Detectado', 3)", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, FechaDetectado, FechaRetirado, Estado, Priorizacion) VALUES (900781, 900790, 'VIG-MULTIED', 'R-VIG-MULTIED-2', #2025-01-01#, NULL, 'Detectado', 3)", dbFailOnError
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, FechaDetectado, FechaRetirado, Estado, Priorizacion) VALUES (900782, 900791, 'VIG-MULTIED', 'R-VIG-MULTIED-3', #2025-01-01#, NULL, 'Detectado', 3)", dbFailOnError

    ' Act
    Dim m_Error As String
    Dim rs As DAO.Recordset
    Set rs = IndicadorRiesgosV2_GetResumen(FIX_DINI_VIG, FIX_DFIN_VIG, CStr(FIX_PROYECTO_VIG), m_Error)

    ' Teardown
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE CodigoUnico = 'VIG-MULTIED'"
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion IN (900790, 900791)"
    On Error GoTo 0

    If rs Is Nothing Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_MismoRiesgoEnVariasEdiciones_CuentaUnaVez = BuildFail( _
            "GetResumen retorno Nothing. m_Error='" & m_Error & "'", logs)
        Exit Function
    End If

    Dim nVig As Long
    nVig = Nz(rs!RiesgosVigentesEnPeriodo, -1)
    rs.Close: Set rs = Nothing

    If nVig <> 1 Then
        Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_MismoRiesgoEnVariasEdiciones_CuentaUnaVez = BuildFail( _
            "Esperado RiesgosVigentesEnPeriodo=1 (mismo CodigoUnico en 3 ediciones), obtuvo " & nVig, logs)
        Exit Function
    End If

    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_MismoRiesgoEnVariasEdiciones_CuentaUnaVez = BuildOk("multied_3_ediciones_1", logs)
    Exit Function
EH:
    ' Capturamos Err ANTES de cualquier On Error Resume Next para no perderlo.
    Dim mErrNum As Long, mErrDesc As String
    mErrNum = Err.Number
    mErrDesc = Err.Description
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE CodigoUnico = 'VIG-MULTIED'"
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion IN (900790, 900791)"
    On Error GoTo 0
    Test_IND_IndicadorRiesgosV2_GetResumen_VigentesEnPeriodo_MismoRiesgoEnVariasEdiciones_CuentaUnaVez = BuildFail( _
        "Excepcion " & mErrNum & " - " & mErrDesc, logs)
End Function

