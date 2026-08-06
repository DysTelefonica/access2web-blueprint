Attribute VB_Name = "Test_RiesgoValoracionHelper"
Option Compare Database
Option Explicit

' ============================================================
' Test_RiesgoValoracionHelper - TDD atoms for modRiesgoValoracionHelper
'
' Helper: CalcularValoracion
'   Signature: Public Function CalcularValoracion( _
'                 ByRef p_IDRiesgo As String
'                 Optional ByRef p_Error As String) As String
'
' Helper resuelve la Valoracion de un riesgo cruzando TbRiesgos (ImpactoGlobal
' + Vulnerabilidad) contra la matriz TbRiesgosValoracion.
'
' Architecture:
'   - DB-driven helper: helper acepta `db` ByRef opcional; átomos TDD
'     inyectan sandbox via Test_Fixtures.GetTestDb().
'   - Schema-first: TbRiesgosValoracion tiene columnas Impacto, Vulnerabilidad
'     y Valoracion como dbText (size 255). NO son numéricas. La matriz
'     staging tiene 25 filas (5x5) con etiquetas:
'         Impacto ? {Muy Bajo, Bajo, Medio, Alto, Muy Alto}
'         Vulnerabilidad ? {Muy Bajo, Bajo, Medio, Alto, Muy Alto}
'   - El átomo adversarial (atom 4) modifica manualmente la matriz
'     (UPDATE fila existente) y verifica que el helper lee DESDE BD
'     (no calcula en VBA). Esto cierra el contrato: la matriz es la
'     fuente de verdad.
'
' Fixture IDs: 905001-905099 (sin overlap con:
'   - Bloque 2: 900500-900599, 50010+
'   - Bloque 3 / REQ-CAL-07: 901000-901099
'   - Bloque 3 / REQ-CAL-09/10: 902000-902099
'   - Bloque 3 / REQ-CAL-12: 904000-904099
'   - Bloque 4 / REQ-CAL-04: helper puro, sin fixtures DB).
'
' Skill: access-vba-tdd v2.4.3
' SDD:    e2e-form-by-form-2026-06-22 - Bloque 4 - REQ-CAL-05
' ============================================================

' --- Module-level constants for REQ-CAL-05 fixtures ---
'     Usamos un único grafo padre compartido (Expediente, Proyecto, Edición)
'     y solo variamos el riesgo hijo por átomo.
Private Const FIX_ID_EXPEDIENTE As Long = 905001
Private Const FIX_ID_PROYECTO   As Long = 905002
Private Const FIX_ID_EDICION    As Long = 905003

Private Const FIX_ID_RIESGO_HAPPY As Long = 905010
Private Const FIX_ID_RIESGO_SAD   As Long = 905011
Private Const FIX_ID_RIESGO_EDGE  As Long = 905012
Private Const FIX_ID_RIESGO_ADV   As Long = 905013

' --- JSON helpers wrapper (project convention) ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' --- EnsureTestConfigLoaded delegation ---
Private Function EnsureTestConfigLoaded(ByRef p_Error As String) As Boolean
    EnsureTestConfigLoaded = Test_Helper.EnsureTestConfigLoaded(p_Error)
End Function

' --- Semilla del grafo padre + un riesgo con valores (ImpactoGlobal,
'     Vulnerabilidad) parametrizados. El grafo padre (Expediente, Proyecto,
'     Edición) se siembra una vez; el riesgo se reemplaza por átomo.
Private Sub SeedR5ParentGraph(ByRef db As DAO.Database)
    On Error GoTo EH_Seed

    ' Limpieza previa (FK-inversa) — solo nuestras fixtures
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo IN (" & _
               FIX_ID_RIESGO_HAPPY & "," & FIX_ID_RIESGO_SAD & "," & _
               FIX_ID_RIESGO_EDGE & "," & FIX_ID_RIESGO_ADV & ")", _
               dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_ID_PROYECTO, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_ID_EXPEDIENTE, dbFailOnError

    ' 1. Expediente (padre raíz)
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
               "VALUES (" & FIX_ID_EXPEDIENTE & ", 'TEST05', 'Fixture REQ-CAL-05', 'Test', 1)", _
               dbFailOnError

    ' 2. Proyecto (hijo de Expediente)
    db.Execute "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto) " & _
               "VALUES (" & FIX_ID_PROYECTO & ", " & FIX_ID_EXPEDIENTE & ", 'TESTPROJ05')", _
               dbFailOnError

    ' 3. Edición (hijo de Proyecto)
    db.Execute "INSERT INTO TbProyectosEdiciones (IDProyecto, IDEdicion, Edicion, Elaborado) " & _
               "VALUES (" & FIX_ID_PROYECTO & ", " & FIX_ID_EDICION & ", 1, 'test_user')", _
               dbFailOnError

    Exit Sub

EH_Seed:
    Dim eN As Long: eN = Err.Number
    Dim ed As String: ed = Err.description
    Err.Raise eN, "SeedR5ParentGraph", "Seed failed: " & eN & " - " & ed
End Sub

' --- Siembra de un riesgo hijo con (ImpactoGlobal, Vulnerabilidad)
'     parametrizados. Schema-first: CodigoUnico + CodigoRiesgo required.
Private Sub SeedR5Riesgo(ByRef db As DAO.Database, _
                          ByVal p_IDRiesgo As Long, _
                          ByVal p_Codigo As String, _
                          ByVal p_ImpactoGlobal As String, _
                          ByVal p_Vulnerabilidad As String)
    Dim m_SQL As String
    m_SQL = "INSERT INTO TbRiesgos " & _
            "(IDRiesgo, IDEdicion, CodigoUnico, CodigoRiesgo, ImpactoGlobal, Vulnerabilidad) " & _
            "VALUES (" & p_IDRiesgo & ", " & FIX_ID_EDICION & ", " & _
            "'UNICO-R5-" & p_IDRiesgo & "', " & _
            "'" & Test_Helper.SqlStr(p_Codigo) & "', " & _
            "'" & Test_Helper.SqlStr(p_ImpactoGlobal) & "', " & _
            "'" & Test_Helper.SqlStr(p_Vulnerabilidad) & "')"
    db.Execute m_SQL, dbFailOnError
End Sub

' --- Teardown FK-inverso: limpia solo nuestras fixtures ---
Private Sub TeardownR5(ByRef db As DAO.Database)
    On Error Resume Next
    db.Execute "DELETE FROM TbRiesgos WHERE IDRiesgo IN (" & _
               FIX_ID_RIESGO_HAPPY & "," & FIX_ID_RIESGO_SAD & "," & _
               FIX_ID_RIESGO_EDGE & "," & FIX_ID_RIESGO_ADV & ")", _
               dbFailOnError
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDEdicion=" & FIX_ID_EDICION, dbFailOnError
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & FIX_ID_PROYECTO, dbFailOnError
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & FIX_ID_EXPEDIENTE, dbFailOnError
    On Error GoTo 0
End Sub

' ============================================================
' ATOM 1 — Happy: ImpactoGlobal=Alto + Vulnerabilidad=Bajo ? "Alto"
' GIVEN sandbox + riesgo con ImpactoGlobal="Alto", Vulnerabilidad="Bajo"
' WHEN  CalcularValoracion(FIX_ID_RIESGO_HAPPY, db, err)
' THEN  retorna "Alto" (matriz), p_Error vacío
'
' Contrato: la matriz TbRiesgosValoracion dice
'   Impacto=Alto, Vulnerabilidad=Bajo ? Valoracion="Alto" (verificado vía
'   dysflow_query_sql el 2026-06-22).
' ============================================================
Public Function Test_RiesgoValoracionHelper_Happy_ImpactoAltoVulnerabilidadBaja_RetornaAlto() As String
    Dim logs(0 To 7) As String
    Test_RiesgoValoracionHelper_Happy_ImpactoAltoVulnerabilidadBaja_RetornaAlto = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedR5ParentGraph + SeedR5Riesgo(" & FIX_ID_RIESGO_HAPPY & _
              ", ImpactoGlobal='Alto', Vulnerabilidad='Bajo')"
    logs(2) = "3. Act: CalcularValoracion"
    logs(3) = "4. Assert: retorno no vacío"
    logs(4) = "5. Assert: retorno = 'Alto' (literal de la matriz)"
    logs(5) = "6. Assert: p_Error vacío"

    Dim cfgErr As String
    If Not EnsureTestConfigLoaded(cfgErr) Then
        Test_RiesgoValoracionHelper_Happy_ImpactoAltoVulnerabilidadBaja_RetornaAlto = _
            BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoValoracionHelper_Happy_ImpactoAltoVulnerabilidadBaja_RetornaAlto = _
            BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    SeedR5ParentGraph db
    SeedR5Riesgo db, FIX_ID_RIESGO_HAPPY, "R-HAPPY", "Alto", "Bajo"

    Dim m_Result As String
    Dim m_Err As String
    m_Result = CalcularValoracion(CStr(FIX_ID_RIESGO_HAPPY), db, m_Err)

    If Len(m_Err) <> 0 Then
        logs(5) = "6. Assert FAIL: p_Error no esperado: " & m_Err
        Test_RiesgoValoracionHelper_Happy_ImpactoAltoVulnerabilidadBaja_RetornaAlto = _
            BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If Len(m_Result) = 0 Then
        logs(4) = "5. Assert FAIL: retorno vacío, esperaba 'Alto'"
        Test_RiesgoValoracionHelper_Happy_ImpactoAltoVulnerabilidadBaja_RetornaAlto = _
            BuildFail("retorno vacío, esperaba 'Alto'", logs)
        GoTo Teardown
    End If
    If m_Result <> "Alto" Then
        logs(4) = "5. Assert FAIL: esperaba 'Alto', obtuvo: '" & m_Result & "'"
        Test_RiesgoValoracionHelper_Happy_ImpactoAltoVulnerabilidadBaja_RetornaAlto = _
            BuildFail("esperaba 'Alto', obtuvo: '" & m_Result & "'", logs)
        GoTo Teardown
    End If

    logs(6) = "7. Assert PASS: 'Alto' desde la matriz (Impacto=Alto, Vulnerabilidad=Bajo)"
    Test_RiesgoValoracionHelper_Happy_ImpactoAltoVulnerabilidadBaja_RetornaAlto = _
        BuildOk(m_Result, logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not db Is Nothing Then TeardownR5 db
    Set db = Nothing
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_RiesgoValoracionHelper_Happy_ImpactoAltoVulnerabilidadBaja_RetornaAlto = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 2 — Sad: ImpactoGlobal vacío ? "" + p_Error poblado
' GIVEN sandbox + riesgo con ImpactoGlobal=""
' WHEN  CalcularValoracion(FIX_ID_RIESGO_SAD, db, err)
' THEN  retorna "", p_Error menciona "ImpactoGlobal"
' ============================================================
Public Function Test_RiesgoValoracionHelper_Sad_RiesgoSinImpactoGlobal_RetornaVacioYPopulaError() As String
    Dim logs(0 To 6) As String
    Test_RiesgoValoracionHelper_Sad_RiesgoSinImpactoGlobal_RetornaVacioYPopulaError = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedR5ParentGraph + SeedR5Riesgo(" & FIX_ID_RIESGO_SAD & _
              ", ImpactoGlobal='', Vulnerabilidad='Bajo')"
    logs(2) = "3. Act: CalcularValoracion"
    logs(3) = "4. Assert: retorno = '' (cadena vacía)"
    logs(4) = "5. Assert: p_Error menciona 'ImpactoGlobal'"

    Dim cfgErr As String
    If Not EnsureTestConfigLoaded(cfgErr) Then
        Test_RiesgoValoracionHelper_Sad_RiesgoSinImpactoGlobal_RetornaVacioYPopulaError = _
            BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoValoracionHelper_Sad_RiesgoSinImpactoGlobal_RetornaVacioYPopulaError = _
            BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    SeedR5ParentGraph db
    SeedR5Riesgo db, FIX_ID_RIESGO_SAD, "R-SAD", "", "Bajo"

    Dim m_Result As String
    Dim m_Err As String
    m_Result = CalcularValoracion(CStr(FIX_ID_RIESGO_SAD), db, m_Err)

    If Len(m_Result) <> 0 Then
        logs(3) = "4. Assert FAIL: esperaba '', obtuvo: '" & m_Result & "'"
        Test_RiesgoValoracionHelper_Sad_RiesgoSinImpactoGlobal_RetornaVacioYPopulaError = _
            BuildFail("esperaba '', obtuvo: '" & m_Result & "'", logs)
        GoTo Teardown
    End If

    If Len(m_Err) = 0 Then
        logs(4) = "5. Assert FAIL: p_Error vacío, esperaba mensaje sobre ImpactoGlobal"
        Test_RiesgoValoracionHelper_Sad_RiesgoSinImpactoGlobal_RetornaVacioYPopulaError = _
            BuildFail("p_Error vacío, esperaba mensaje sobre ImpactoGlobal", logs)
        GoTo Teardown
    End If

    If InStr(1, m_Err, "ImpactoGlobal", vbTextCompare) = 0 Then
        logs(4) = "5. Assert FAIL: p_Error no menciona 'ImpactoGlobal': " & m_Err
        Test_RiesgoValoracionHelper_Sad_RiesgoSinImpactoGlobal_RetornaVacioYPopulaError = _
            BuildFail("p_Error no menciona 'ImpactoGlobal': " & m_Err, logs)
        GoTo Teardown
    End If

    logs(5) = "6. Assert PASS: '' + p_Error describe el problema"
    Test_RiesgoValoracionHelper_Sad_RiesgoSinImpactoGlobal_RetornaVacioYPopulaError = _
        BuildOk("sad_sin_impacto_global", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not db Is Nothing Then TeardownR5 db
    Set db = Nothing
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_RiesgoValoracionHelper_Sad_RiesgoSinImpactoGlobal_RetornaVacioYPopulaError = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 3 — Edge: Vulnerabilidad fuera de rango ? "" + p_Error
' GIVEN sandbox + riesgo con ImpactoGlobal="Muy Bajo",
'       Vulnerabilidad="ZZZ_Invalido" (fuera de la matriz 5x5)
' WHEN  CalcularValoracion(FIX_ID_RIESGO_EDGE, db, err)
' THEN  retorna "", p_Error menciona "no se encontró Valoracion"
' ============================================================
Public Function Test_RiesgoValoracionHelper_Edge_CombinacionInexistenteEnMatriz_RetornaVacioYPopulaError() As String
    Dim logs(0 To 6) As String
    Test_RiesgoValoracionHelper_Edge_CombinacionInexistenteEnMatriz_RetornaVacioYPopulaError = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedR5ParentGraph + SeedR5Riesgo(" & FIX_ID_RIESGO_EDGE & _
              ", ImpactoGlobal='Muy Bajo', Vulnerabilidad='ZZZ_Invalido')"
    logs(2) = "3. Act: CalcularValoracion (combinación fuera de matriz)"
    logs(3) = "4. Assert: retorno = '' (cadena vacía)"
    logs(4) = "5. Assert: p_Error menciona 'no se encontró Valoracion'"

    Dim cfgErr As String
    If Not EnsureTestConfigLoaded(cfgErr) Then
        Test_RiesgoValoracionHelper_Edge_CombinacionInexistenteEnMatriz_RetornaVacioYPopulaError = _
            BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoValoracionHelper_Edge_CombinacionInexistenteEnMatriz_RetornaVacioYPopulaError = _
            BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    SeedR5ParentGraph db
    SeedR5Riesgo db, FIX_ID_RIESGO_EDGE, "R-EDGE", "Muy Bajo", "ZZZ_Invalido"

    Dim m_Result As String
    Dim m_Err As String
    m_Result = CalcularValoracion(CStr(FIX_ID_RIESGO_EDGE), db, m_Err)

    If Len(m_Result) <> 0 Then
        logs(3) = "4. Assert FAIL: esperaba '', obtuvo: '" & m_Result & "'"
        Test_RiesgoValoracionHelper_Edge_CombinacionInexistenteEnMatriz_RetornaVacioYPopulaError = _
            BuildFail("esperaba '', obtuvo: '" & m_Result & "'", logs)
        GoTo Teardown
    End If

    If Len(m_Err) = 0 Then
        logs(4) = "5. Assert FAIL: p_Error vacío, esperaba mensaje sobre fila faltante"
        Test_RiesgoValoracionHelper_Edge_CombinacionInexistenteEnMatriz_RetornaVacioYPopulaError = _
            BuildFail("p_Error vacío, esperaba mensaje sobre fila faltante", logs)
        GoTo Teardown
    End If

    If InStr(1, m_Err, "no se encontró Valoracion", vbTextCompare) = 0 Then
        logs(4) = "5. Assert FAIL: p_Error no menciona 'no se encontró Valoracion': " & m_Err
        Test_RiesgoValoracionHelper_Edge_CombinacionInexistenteEnMatriz_RetornaVacioYPopulaError = _
            BuildFail("p_Error no describe fila faltante: " & m_Err, logs)
        GoTo Teardown
    End If

    logs(5) = "6. Assert PASS: '' + p_Error describe fila faltante"
    Test_RiesgoValoracionHelper_Edge_CombinacionInexistenteEnMatriz_RetornaVacioYPopulaError = _
        BuildOk("edge_combinacion_inexistente", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    If Not db Is Nothing Then TeardownR5 db
    Set db = Nothing
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    Test_RiesgoValoracionHelper_Edge_CombinacionInexistenteEnMatriz_RetornaVacioYPopulaError = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
    Resume Teardown
End Function

' ============================================================
' ATOM 4 — Adversarial: matriz modificada manualmente ?
'                       helper lee "Personalizado" (no "Alto")
' GIVEN sandbox + riesgo con ImpactoGlobal="Alto", Vulnerabilidad="Bajo"
'       Y la fila (Alto, Bajo) de TbRiesgosValoracion modificada a
'       Valoracion="Personalizado" (override del default "Alto")
' WHEN  CalcularValoracion(FIX_ID_RIESGO_ADV, db, err)
' THEN  retorna "Personalizado" — esto prueba que el helper lee desde BD
'       (no calcula en VBA). Lock-in de contrato: la matriz es la fuente
'       de verdad.
'
' Teardown: restaura la fila (Alto, Bajo) a su Valoracion original
'           para no contaminar otros tests que dependan del default.
' ============================================================
Public Function Test_RiesgoValoracionHelper_Adversarial_MatrizModificadaManualmente_DetectaDivergencia() As String
    Dim logs(0 To 9) As String
    Test_RiesgoValoracionHelper_Adversarial_MatrizModificadaManualmente_DetectaDivergencia = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    Dim m_OriginalValoracion As String
    m_OriginalValoracion = ""

    logs(0) = "1. Arrange: EnsureTestConfigLoaded"
    logs(1) = "2. Arrange: SeedR5ParentGraph + SeedR5Riesgo(" & FIX_ID_RIESGO_ADV & _
              ", ImpactoGlobal='Alto', Vulnerabilidad='Bajo')"
    logs(2) = "3. Arrange: capturar Valoracion original de la fila (Alto, Bajo)"
    logs(3) = "4. Arrange: UPDATE matriz ? Valoracion='Personalizado' (override manual)"
    logs(4) = "5. Act: CalcularValoracion"
    logs(5) = "6. Assert: retorno = 'Personalizado' (lectura desde BD, no VBA)"
    logs(6) = "7. Teardown: UPDATE matriz ? restaurar Valoracion original"

    Dim cfgErr As String
    If Not EnsureTestConfigLoaded(cfgErr) Then
        Test_RiesgoValoracionHelper_Adversarial_MatrizModificadaManualmente_DetectaDivergencia = _
            BuildFail(cfgErr, logs)
        GoTo Teardown
    End If

    Dim db As DAO.Database
    Dim dbErr As String
    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_RiesgoValoracionHelper_Adversarial_MatrizModificadaManualmente_DetectaDivergencia = _
            BuildFail("GetTestDb Nothing: " & dbErr, logs)
        GoTo Teardown
    End If

    SeedR5ParentGraph db
    SeedR5Riesgo db, FIX_ID_RIESGO_ADV, "R-ADV", "Alto", "Bajo"

    ' --- Capturar Valoracion original de la fila (Alto, Bajo) ---
    Dim m_SnapshotRS As DAO.Recordset
    Set m_SnapshotRS = db.OpenRecordset( _
        "SELECT Valoracion FROM TbRiesgosValoracion " & _
        "WHERE Impacto='Alto' AND Vulnerabilidad='Bajo'", _
        dbOpenSnapshot)
    If m_SnapshotRS.EOF Then
        m_SnapshotRS.Close
        Set m_SnapshotRS = Nothing
        Test_RiesgoValoracionHelper_Adversarial_MatrizModificadaManualmente_DetectaDivergencia = _
            BuildFail("pre-check: no existe fila (Alto, Bajo) en la matriz", logs)
        GoTo Teardown
    End If
    m_OriginalValoracion = CStr(Nz(m_SnapshotRS.fields("Valoracion").value, ""))
    m_SnapshotRS.Close
    Set m_SnapshotRS = Nothing
    logs(2) = "3. Snapshot: Valoracion original='" & m_OriginalValoracion & "'"

    ' --- Override: UPDATE la fila (Alto, Bajo) a 'Personalizado' ---
    db.Execute "UPDATE TbRiesgosValoracion " & _
               "SET Valoracion='Personalizado' " & _
               "WHERE Impacto='Alto' AND Vulnerabilidad='Bajo'", _
               dbFailOnError
    logs(3) = "4. Override aplicado: Valoracion='Personalizado'"

    ' --- Act: helper debe leer DESDE BD ---
    Dim m_Result As String
    Dim m_Err As String
    m_Result = CalcularValoracion(CStr(FIX_ID_RIESGO_ADV), db, m_Err)

    ' --- Assert: retorno = 'Personalizado' ---
    If Len(m_Err) <> 0 Then
        logs(5) = "6. Assert FAIL: p_Error no esperado: " & m_Err
        Test_RiesgoValoracionHelper_Adversarial_MatrizModificadaManualmente_DetectaDivergencia = _
            BuildFail("p_Error no esperado: " & m_Err, logs)
        GoTo Teardown
    End If
    If m_Result <> "Personalizado" Then
        logs(5) = "6. Assert FAIL: esperaba 'Personalizado' (override BD), obtuvo: '" & m_Result & "'"
        Test_RiesgoValoracionHelper_Adversarial_MatrizModificadaManualmente_DetectaDivergencia = _
            BuildFail("helper NO lee de BD; calculó '" & m_Result & "' en VBA", logs)
        GoTo Teardown
    End If

    logs(6) = "7. Assert PASS: helper lee 'Personalizado' desde BD (no VBA-calculado)"
    logs(7) = "8. Contrato cerrado: matriz TbRiesgosValoracion es la fuente de verdad"
    Test_RiesgoValoracionHelper_Adversarial_MatrizModificadaManualmente_DetectaDivergencia = _
        BuildOk("adversarial_matriz_bd_es_fuente_de_verdad", logs)
    GoTo Teardown

Teardown:
    On Error Resume Next
    ' --- Restaurar Valoracion original (independiente del resultado del test) ---
    If Len(m_OriginalValoracion) > 0 And Not db Is Nothing Then
        db.Execute "UPDATE TbRiesgosValoracion " & _
                   "SET Valoracion='" & Test_Helper.SqlStr(m_OriginalValoracion) & "' " & _
                   "WHERE Impacto='Alto' AND Vulnerabilidad='Bajo'", _
                   dbFailOnError
    End If
    If Not db Is Nothing Then TeardownR5 db
    Set db = Nothing
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Exit Function

HandleError:
    ' Restaurar ANTES de reportar (para no contaminar la matriz aunque
    ' el handler reporte fallo)
    On Error Resume Next
    If Len(m_OriginalValoracion) > 0 And Not db Is Nothing Then
        db.Execute "UPDATE TbRiesgosValoracion " & _
                   "SET Valoracion='" & Test_Helper.SqlStr(m_OriginalValoracion) & "' " & _
                   "WHERE Impacto='Alto' AND Vulnerabilidad='Bajo'", _
                   dbFailOnError
    End If
    If Not db Is Nothing Then TeardownR5 db
    Set db = Nothing
    Test_Helper.ResetTestSession
    On Error GoTo 0
    Test_RiesgoValoracionHelper_Adversarial_MatrizModificadaManualmente_DetectaDivergencia = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
End Function

