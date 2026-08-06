Attribute VB_Name = "Test_RiesgoPriorizacionHelper"
' ============================================================
' Test_RiesgoPriorizacionHelper — TDD Atoms
'
' Skill: access-vba-tdd
' Scope:
'   - ValidarPriorizacion helper contract
'   - GenerarEdicionNuevaAPartirDeAnterior priorizacion copy contract
'
' Assumptions (?contract):
'   - Valid range: 1 <= p_Prioridad <= p_NumRiesgosActivos (inclusive)
'   - Non-numeric p_Prioridad > False
'   - p_Prioridad < 1 or > p_NumRiesgosActivos > False
'   - Whitespace: trimmed before validation
'
' Atoms: 6
' ============================================================
Option Compare Database
Option Explicit

' --- RunAll ---
Public Function Test_RiesgoPriorizacion_RunAll() As String
    Dim logs(0 To 0) As String
    logs(0) = "Runner: Test_RiesgoPriorizacion_RunAll"
    Test_RiesgoPriorizacion_RunAll = BuildJsonOk("ok", logs)
End Function

' ============================================================
' ATOM: Test_ValidarPriorizacion_Happy
' Scenario: happy — valid integer inside range
' Input: p_Prioridad="3", p_NumRiesgosActivos=10
' Expected: True
' ============================================================
Public Function Test_ValidarPriorizacion_Happy() As String
    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: p_Prioridad=""3"", p_NumRiesgosActivos=10"
    logs(1) = "2. Act: ValidarPriorizacion(""3"", 10)"
    logs(2) = "3. Assert: result = True"

    Dim result As Boolean

    On Error GoTo EH

    ' RED phase: modRiesgoPriorizacionHelper does not exist yet.
    ' This line will raise a compile-time or runtime error.
    result = ValidarPriorizacion("3", 10)

    If result <> True Then
        Test_ValidarPriorizacion_Happy = BuildJsonFail( _
            "Expected True but got False", logs)
        Exit Function
    End If

    Test_ValidarPriorizacion_Happy = BuildJsonOk(result, logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_ValidarPriorizacion_Happy = BuildJsonFail( _
        "ValidarPriorizacion raised: " & Err.description, logs)
End Function

' ============================================================
' ATOM: Test_ValidarPriorizacion_Sad
' Scenario: sad — invalid inputs (blank, non-numeric, negative)
' Cases:
'   a) p_Prioridad="" (blank)
'   b) p_Prioridad="abc" (non-numeric)
'   c) p_Prioridad="-1" (negative, outside valid range)
' Expected: all False
' ============================================================
Public Function Test_ValidarPriorizacion_Sad() As String
    Dim logs(0 To 8) As String
    logs(0) = "1. Arrange: sad cases"
    logs(1) = "2. Case a: p_Prioridad="""" (blank)"
    logs(2) = "3. Case b: p_Prioridad=""abc"" (non-numeric)"
    logs(3) = "4. Case c: p_Prioridad=""-1"" (negative)"
    logs(4) = "5. Act: ValidarPriorizacion for each case"
    logs(5) = "6. Assert: all results = False"

    Dim resultA As Boolean
    Dim resultB As Boolean
    Dim resultC As Boolean
    Dim failures(0 To 2) As String
    Dim failureCount As Long
    failureCount = 0

    On Error GoTo EH

    ' Case a: blank
    resultA = ValidarPriorizacion("", 10)
    If resultA <> False Then
        failures(failureCount) = "Case a (blank) returned " & CStr(resultA) & ", expected False"
        failureCount = failureCount + 1
    End If

    ' Case b: non-numeric
    resultB = ValidarPriorizacion("abc", 10)
    If resultB <> False Then
        failures(failureCount) = "Case b (non-numeric) returned " & CStr(resultB) & ", expected False"
        failureCount = failureCount + 1
    End If

    ' Case c: negative integer
    resultC = ValidarPriorizacion("-1", 10)
    If resultC <> False Then
        failures(failureCount) = "Case c (negative) returned " & CStr(resultC) & ", expected False"
        failureCount = failureCount + 1
    End If

    If failureCount > 0 Then
        Dim allFailures As String
        allFailures = Join(failures, " | ")
        logs(6) = "7. Failures: " & allFailures
        Test_ValidarPriorizacion_Sad = BuildJsonFail(allFailures, logs)
        Exit Function
    End If

    Test_ValidarPriorizacion_Sad = BuildJsonOk(failureCount, logs)
    Exit Function

EH:
    logs(5) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_ValidarPriorizacion_Sad = BuildJsonFail( _
        "ValidarPriorizacion raised: " & Err.description, logs)
End Function

' ============================================================
' ATOM: Test_ValidarPriorizacion_Edge
' Scenario: edge — boundary and near-boundary values
' Cases:
'   a) p_Prioridad=" 3 " (whitespace padded — trimmed)
'   b) p_Prioridad="3.5" (decimal — non-integer)
'   c) p_NumRiesgosActivos=1 (minimum valid N, p_Prioridad="1")
'   d) p_NumRiesgosActivos=0 (degenerate — no active risks)
' Expected: a,b,d > False; c > True
' ============================================================
Public Function Test_ValidarPriorizacion_Edge() As String
    Dim logs(0 To 9) As String
    logs(0) = "1. Arrange: edge cases"
    logs(1) = "2. Case a: p_Prioridad="" 3 "" (whitespace)"
    logs(2) = "3. Case b: p_Prioridad=""3.5"" (decimal)"
    logs(3) = "4. Case c: p_NumRiesgosActivos=1, p_Prioridad=""1"" (min boundary)"
    logs(4) = "5. Case d: p_NumRiesgosActivos=0 (degenerate)"
    logs(5) = "6. Act: ValidarPriorizacion for each case"
    logs(6) = "7. Assert: a=False, b=False, c=True, d=False"

    Dim resultA As Boolean
    Dim resultB As Boolean
    Dim resultC As Boolean
    Dim resultD As Boolean
    Dim failures(0 To 3) As String
    Dim failureCount As Long
    failureCount = 0

    On Error GoTo EH

    ' Case a: whitespace — trimmed before validation
    resultA = ValidarPriorizacion(" 3 ", 10)
    If resultA <> False Then
        failures(failureCount) = "Case a (whitespace) returned " & CStr(resultA) & ", expected False"
        failureCount = failureCount + 1
    End If

    ' Case b: decimal — not a valid integer
    resultB = ValidarPriorizacion("3.5", 10)
    If resultB <> False Then
        failures(failureCount) = "Case b (decimal) returned " & CStr(resultB) & ", expected False"
        failureCount = failureCount + 1
    End If

    ' Case c: boundary — p_NumRiesgosActivos=1, valid priority "1"
    resultC = ValidarPriorizacion("1", 1)
    If resultC <> True Then
        failures(failureCount) = "Case c (min boundary) returned " & CStr(resultC) & ", expected True"
        failureCount = failureCount + 1
    End If

    ' Case d: degenerate — no active risks
    resultD = ValidarPriorizacion("1", 0)
    If resultD <> False Then
        failures(failureCount) = "Case d (zero risks) returned " & CStr(resultD) & ", expected False"
        failureCount = failureCount + 1
    End If

    If failureCount > 0 Then
        Dim allFailures As String
        allFailures = Join(failures, " | ")
        logs(7) = "8. Failures: " & allFailures
        Test_ValidarPriorizacion_Edge = BuildJsonFail(allFailures, logs)
        Exit Function
    End If

    Test_ValidarPriorizacion_Edge = BuildJsonOk(failureCount, logs)
    Exit Function

EH:
    logs(6) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_ValidarPriorizacion_Edge = BuildJsonFail( _
        "ValidarPriorizacion raised: " & Err.description, logs)
End Function

' ============================================================
' ATOM: Test_ValidarPriorizacion_Adversarial
' Scenario: adversarial — hostile inputs
' Cases:
'   a) p_Prioridad contains special chars (e.g. "1; DROP TABLE")
'   b) p_Prioridad is Null (VarType 10)
'   c) p_NumRiesgosActivos=-5 (negative active count)
' Expected: all False
' ============================================================
Public Function Test_ValidarPriorizacion_Adversarial() As String
    Dim logs(0 To 8) As String
    logs(0) = "1. Arrange: adversarial cases"
    logs(1) = "2. Case a: p_Prioridad=""1; DROP TABLE tbRiesgos"" (special chars)"
    logs(2) = "3. Case b: p_Prioridad = Null (VarType 10)"
    logs(3) = "4. Case c: p_NumRiesgosActivos=-5 (negative count)"
    logs(4) = "5. Act: ValidarPriorizacion for each case"
    logs(5) = "6. Assert: all results = False"

    Dim resultA As Boolean
    Dim resultB As Boolean
    Dim resultC As Boolean
    Dim p_PrioridadB As Variant
    Dim failures(0 To 2) As String
    Dim failureCount As Long
    failureCount = 0

    On Error GoTo EH

    ' Case a: special characters (SQL injection attempt)
    resultA = ValidarPriorizacion("1; DROP TABLE tbRiesgos", 10)
    If resultA <> False Then
        failures(failureCount) = "Case a (special chars) returned " & CStr(resultA) & ", expected False"
        failureCount = failureCount + 1
    End If

    ' Case b: Null as p_Prioridad
    p_PrioridadB = Null
    resultB = ValidarPriorizacion(p_PrioridadB, 10)
    If resultB <> False Then
        failures(failureCount) = "Case b (Null) returned " & CStr(resultB) & ", expected False"
        failureCount = failureCount + 1
    End If

    ' Case c: negative p_NumRiesgosActivos
    resultC = ValidarPriorizacion("1", -5)
    If resultC <> False Then
        failures(failureCount) = "Case c (negative count) returned " & CStr(resultC) & ", expected False"
        failureCount = failureCount + 1
    End If

    If failureCount > 0 Then
        Dim allFailures As String
        allFailures = Join(failures, " | ")
        logs(6) = "7. Failures: " & allFailures
        Test_ValidarPriorizacion_Adversarial = BuildJsonFail(allFailures, logs)
        Exit Function
    End If

    Test_ValidarPriorizacion_Adversarial = BuildJsonOk(failureCount, logs)
    Exit Function

EH:
    logs(5) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_ValidarPriorizacion_Adversarial = BuildJsonFail( _
        "ValidarPriorizacion raised: " & Err.description, logs)
End Function

' ============================================================
' ATOM 5: Escenario 5 del contrato (Punto 03 acta 25/06)
' "La priorizacion no arrastra valor de la edicion anterior"
' Test: al generar una nueva edicion para un proyecto sin biblioteca,
' GenerarEdicionNuevaAPartirDeAnterior debe dejar vacia la Priorizacion
' del riesgo copiado, independientemente del valor de la edicion anterior.
' ============================================================
Public Function Test_GenerarEdicionNueva_NoBiblioteca_PriorizacionNoArrastraDeAnterior() As String
    Dim logs(0 To 10) As String
    Dim db As DAO.Database
    Dim dbErr As String
    Dim m_IDExpediente As Long
    Dim m_IDProyecto As Long
    Dim m_IDEdicionOrigen As Long
    Dim m_IDEdicionNueva As Long
    Dim m_IDRiesgoOrigen As Long
    Dim m_IDRiesgoCopia As Long
    Dim m_PriorizacionOrigen As String
    Dim m_PriorizacionCopia As String
    Dim rcd As DAO.Recordset
    Dim m_EdicionNueva As Edicion
    Dim errGenerar As String
    Dim countAfter As Long

    On Error GoTo EH

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_GenerarEdicionNueva_NoBiblioteca_PriorizacionNoArrastraDeAnterior = BuildJsonFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    m_IDExpediente = 950101
    m_IDProyecto = 950102
    m_IDEdicionOrigen = 950103
    m_IDRiesgoOrigen = 950104
    m_PriorizacionOrigen = "5"

    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion IN (SELECT IDEdicion FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto & ")"
    db.Execute "DELETE FROM TbRiesgos WHERE IDEdicion IN (SELECT IDEdicion FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto & ") OR IDRiesgo=" & m_IDRiesgoOrigen
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & m_IDProyecto
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & m_IDExpediente
    On Error GoTo EH

    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
        "VALUES (" & m_IDExpediente & ", 'TESTP03N', 'Test Punto 03 no biblioteca', 'Test', 1)"
    db.Execute "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto, RequiereRiesgoDeBiblioteca) " & _
        "VALUES (" & m_IDProyecto & ", " & m_IDExpediente & ", 'TESTP03NPROY', 'No')"
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado, FechaPublicacion) " & _
        "VALUES (" & m_IDEdicionOrigen & ", " & m_IDProyecto & ", 1, 'TESTUSER', #2026-07-09#)"
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoRiesgo, Descripcion, Estado, Priorizacion, CodigoUnico, FechaDetectado, DetectadoPor, EntidadDetecta, Plazo, Calidad, Coste, ImpactoGlobal, Vulnerabilidad, Valoracion, Mitigacion, Contingencia, RequierePlanContingencia) " & _
        "VALUES (" & m_IDRiesgoOrigen & ", " & m_IDEdicionOrigen & ", 'TESTP03N', 'Test Punto 03 no biblioteca', 'Detectado', '" & m_PriorizacionOrigen & "', 'UNICO-P03-NOBIB', #" & Format$(Now, "yyyy-mm-dd") & "#, 'TESTUSER', 'TEST', 'Medio', 'Medio', 'Medio', 'Medio', 'Medio', 'Medio', 'Reducir', 'Sí', 'Sí')"

    logs(0) = "1. Arrange: proyecto sin biblioteca con edicion publicada y riesgo Priorizacion='" & m_PriorizacionOrigen & "'"
    logs(1) = "2. Act: GenerarEdicionNuevaAPartirDeAnterior sobre IDEdicion=" & m_IDEdicionOrigen
    logs(2) = "3. Assert: la nueva edicion contiene exactamente un riesgo copiado"
    logs(3) = "4. Assert: Priorizacion copiada debe quedar vacia para proyecto sin biblioteca"

    Set m_EdicionNueva = GenerarEdicionNuevaAPartirDeAnterior(CStr(m_IDEdicionOrigen), , errGenerar)
    If errGenerar <> "" Then
        Test_GenerarEdicionNueva_NoBiblioteca_PriorizacionNoArrastraDeAnterior = BuildJsonFail("GenerarEdicionNuevaAPartirDeAnterior error: " & errGenerar, logs)
        Exit Function
    End If
    If m_EdicionNueva Is Nothing Then
        Test_GenerarEdicionNueva_NoBiblioteca_PriorizacionNoArrastraDeAnterior = BuildJsonFail("GenerarEdicionNuevaAPartirDeAnterior retorno Nothing", logs)
        Exit Function
    End If
    m_IDEdicionNueva = CLng(m_EdicionNueva.IDEdicion)

    Set rcd = db.OpenRecordset("SELECT COUNT(*) AS Cnt FROM TbRiesgos WHERE IDEdicion=" & m_IDEdicionNueva, dbOpenSnapshot)
    countAfter = CLng(rcd.Fields("Cnt").Value)
    rcd.Close
    Set rcd = Nothing
    If countAfter <> 1 Then
        logs(4) = "3. Assert FAIL: riesgos copiados=" & CStr(countAfter) & " (esperado 1)"
        Test_GenerarEdicionNueva_NoBiblioteca_PriorizacionNoArrastraDeAnterior = BuildJsonFail("Riesgos copiados=" & CStr(countAfter) & ", esperado 1", logs)
        Exit Function
    End If

    Set rcd = db.OpenRecordset("SELECT TOP 1 IDRiesgo, Priorizacion FROM TbRiesgos WHERE IDEdicion=" & m_IDEdicionNueva, dbOpenSnapshot)
    If rcd.EOF Then
        Test_GenerarEdicionNueva_NoBiblioteca_PriorizacionNoArrastraDeAnterior = BuildJsonFail("No se encontro el riesgo copiado", logs)
        Exit Function
    End If
    m_IDRiesgoCopia = CLng(rcd.Fields("IDRiesgo").Value)
    m_PriorizacionCopia = Nz(rcd.Fields("Priorizacion").Value, "")
    rcd.Close
    Set rcd = Nothing

    logs(4) = "4. Assert: riesgo copiado (ID=" & m_IDRiesgoCopia & ") tiene Priorizacion='" & m_PriorizacionCopia & "' (esperado vacio)"
    If m_PriorizacionCopia <> "" Then
        logs(5) = "4. Assert FAIL: Priorizacion NO es vacia. Con la logica anterior del caller, este atom queda rojo."
        Test_GenerarEdicionNueva_NoBiblioteca_PriorizacionNoArrastraDeAnterior = BuildJsonFail("Priorizacion copiada='" & m_PriorizacionCopia & "', deberia ser vacia", logs)
        Exit Function
    End If

    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion IN (SELECT IDEdicion FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto & ")"
    db.Execute "DELETE FROM TbRiesgos WHERE IDEdicion IN (SELECT IDEdicion FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto & ") OR IDRiesgo=" & m_IDRiesgoOrigen
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & m_IDProyecto
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & m_IDExpediente
    On Error GoTo 0
    Set db = Nothing

    logs(5) = "5. Assert PASS: Priorizacion vacia en el riesgo copiado por GenerarEdicionNuevaAPartirDeAnterior"
    Test_GenerarEdicionNueva_NoBiblioteca_PriorizacionNoArrastraDeAnterior = BuildJsonOk("priorizacion_no_arrastra", logs)
    Exit Function

EH:
    Dim em As String
    em = "Test_GenerarEdicionNueva_NoBiblioteca_PriorizacionNoArrastraDeAnterior fallo: " & Err.description
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion IN (SELECT IDEdicion FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto & ")"
        db.Execute "DELETE FROM TbRiesgos WHERE IDEdicion IN (SELECT IDEdicion FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto & ") OR IDRiesgo=" & m_IDRiesgoOrigen
        db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto
        db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & m_IDProyecto
        db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & m_IDExpediente
    End If
    On Error GoTo 0
    Test_GenerarEdicionNueva_NoBiblioteca_PriorizacionNoArrastraDeAnterior = BuildJsonFail(em, logs)
End Function

' ============================================================
' ATOM 6: Rama opuesta minima — proyecto con biblioteca.
' Si el proyecto requiere riesgo de biblioteca, el caller debe permitir que
' CopiarRiesgo conserve Priorizacion.
' ============================================================
Public Function Test_GenerarEdicionNueva_Biblioteca_PriorizacionSeConserva() As String
    Dim logs(0 To 10) As String
    Dim db As DAO.Database
    Dim dbErr As String
    Dim m_IDExpediente As Long
    Dim m_IDProyecto As Long
    Dim m_IDEdicionOrigen As Long
    Dim m_IDEdicionNueva As Long
    Dim m_IDRiesgoOrigen As Long
    Dim m_PriorizacionOrigen As String
    Dim m_PriorizacionCopia As String
    Dim rcd As DAO.Recordset
    Dim m_EdicionNueva As Edicion
    Dim errGenerar As String
    Dim countAfter As Long

    On Error GoTo EH

    Set db = Test_Fixtures.GetTestDb(dbErr)
    If db Is Nothing Then
        Test_GenerarEdicionNueva_Biblioteca_PriorizacionSeConserva = BuildJsonFail("GetTestDb Nothing: " & dbErr, logs)
        Exit Function
    End If

    m_IDExpediente = 950201
    m_IDProyecto = 950202
    m_IDEdicionOrigen = 950203
    m_IDRiesgoOrigen = 950204
    m_PriorizacionOrigen = "7"

    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion IN (SELECT IDEdicion FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto & ")"
    db.Execute "DELETE FROM TbRiesgos WHERE IDEdicion IN (SELECT IDEdicion FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto & ") OR IDRiesgo=" & m_IDRiesgoOrigen
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & m_IDProyecto
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & m_IDExpediente
    On Error GoTo EH

    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico, Titulo, Tipo, IDEstado) " & _
        "VALUES (" & m_IDExpediente & ", 'TESTP03B', 'Test Punto 03 biblioteca', 'Test', 1)"
    db.Execute "INSERT INTO TbProyectos (IDProyecto, IDExpediente, Proyecto, RequiereRiesgoDeBiblioteca) " & _
        "VALUES (" & m_IDProyecto & ", " & m_IDExpediente & ", 'TESTP03BPROY', 'Sí')"
    db.Execute "INSERT INTO TbProyectosEdiciones (IDEdicion, IDProyecto, Edicion, Elaborado, FechaPublicacion) " & _
        "VALUES (" & m_IDEdicionOrigen & ", " & m_IDProyecto & ", 1, 'TESTUSER', #2026-07-09#)"
    db.Execute "INSERT INTO TbRiesgos (IDRiesgo, IDEdicion, CodigoRiesgo, Descripcion, Estado, Priorizacion, CodigoUnico, FechaDetectado, DetectadoPor, EntidadDetecta, Plazo, Calidad, Coste, ImpactoGlobal, Vulnerabilidad, Valoracion, Mitigacion, Contingencia, RequierePlanContingencia) " & _
        "VALUES (" & m_IDRiesgoOrigen & ", " & m_IDEdicionOrigen & ", 'TESTP03B', 'Test Punto 03 biblioteca', 'Detectado', '" & m_PriorizacionOrigen & "', 'UNICO-P03-BIB', #" & Format$(Now, "yyyy-mm-dd") & "#, 'TESTUSER', 'TEST', 'Medio', 'Medio', 'Medio', 'Medio', 'Medio', 'Medio', 'Reducir', 'Sí', 'Sí')"

    logs(0) = "1. Arrange: proyecto con biblioteca con riesgo Priorizacion='" & m_PriorizacionOrigen & "'"
    logs(1) = "2. Act: GenerarEdicionNuevaAPartirDeAnterior sobre IDEdicion=" & m_IDEdicionOrigen
    logs(2) = "3. Assert: la nueva edicion contiene exactamente un riesgo copiado"
    logs(3) = "4. Assert: Priorizacion se conserva en la rama biblioteca"

    Set m_EdicionNueva = GenerarEdicionNuevaAPartirDeAnterior(CStr(m_IDEdicionOrigen), , errGenerar)
    If errGenerar <> "" Then
        Test_GenerarEdicionNueva_Biblioteca_PriorizacionSeConserva = BuildJsonFail("GenerarEdicionNuevaAPartirDeAnterior error: " & errGenerar, logs)
        Exit Function
    End If
    If m_EdicionNueva Is Nothing Then
        Test_GenerarEdicionNueva_Biblioteca_PriorizacionSeConserva = BuildJsonFail("GenerarEdicionNuevaAPartirDeAnterior retorno Nothing", logs)
        Exit Function
    End If
    m_IDEdicionNueva = CLng(m_EdicionNueva.IDEdicion)

    Set rcd = db.OpenRecordset("SELECT COUNT(*) AS Cnt FROM TbRiesgos WHERE IDEdicion=" & m_IDEdicionNueva, dbOpenSnapshot)
    countAfter = CLng(rcd.Fields("Cnt").Value)
    rcd.Close
    Set rcd = Nothing
    If countAfter <> 1 Then
        logs(4) = "3. Assert FAIL: riesgos copiados=" & CStr(countAfter) & " (esperado 1)"
        Test_GenerarEdicionNueva_Biblioteca_PriorizacionSeConserva = BuildJsonFail("Riesgos copiados=" & CStr(countAfter) & ", esperado 1", logs)
        Exit Function
    End If

    Set rcd = db.OpenRecordset("SELECT TOP 1 Priorizacion FROM TbRiesgos WHERE IDEdicion=" & m_IDEdicionNueva, dbOpenSnapshot)
    If rcd.EOF Then
        Test_GenerarEdicionNueva_Biblioteca_PriorizacionSeConserva = BuildJsonFail("No se encontro el riesgo copiado", logs)
        Exit Function
    End If
    m_PriorizacionCopia = Nz(rcd.Fields("Priorizacion").Value, "")
    rcd.Close
    Set rcd = Nothing

    If m_PriorizacionCopia <> m_PriorizacionOrigen Then
        logs(4) = "4. Assert FAIL: Priorizacion='" & m_PriorizacionCopia & "' (esperado '" & m_PriorizacionOrigen & "')"
        Test_GenerarEdicionNueva_Biblioteca_PriorizacionSeConserva = BuildJsonFail("Priorizacion biblioteca no conservada", logs)
        Exit Function
    End If

    On Error Resume Next
    db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion IN (SELECT IDEdicion FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto & ")"
    db.Execute "DELETE FROM TbRiesgos WHERE IDEdicion IN (SELECT IDEdicion FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto & ") OR IDRiesgo=" & m_IDRiesgoOrigen
    db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto
    db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & m_IDProyecto
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & m_IDExpediente
    On Error GoTo 0
    Set db = Nothing

    logs(4) = "5. Assert PASS: Priorizacion conservada para proyecto con biblioteca"
    Test_GenerarEdicionNueva_Biblioteca_PriorizacionSeConserva = BuildJsonOk("priorizacion_biblioteca_conserva", logs)
    Exit Function

EH:
    Dim em As String
    em = "Test_GenerarEdicionNueva_Biblioteca_PriorizacionSeConserva fallo: " & Err.description
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    If Not db Is Nothing Then
        db.Execute "DELETE FROM TbProyectosEdicionesSuministradores WHERE IDEdicion IN (SELECT IDEdicion FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto & ")"
        db.Execute "DELETE FROM TbRiesgos WHERE IDEdicion IN (SELECT IDEdicion FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto & ") OR IDRiesgo=" & m_IDRiesgoOrigen
        db.Execute "DELETE FROM TbProyectosEdiciones WHERE IDProyecto=" & m_IDProyecto
        db.Execute "DELETE FROM TbProyectos WHERE IDProyecto=" & m_IDProyecto
        db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente=" & m_IDExpediente
    End If
    On Error GoTo 0
    Test_GenerarEdicionNueva_Biblioteca_PriorizacionSeConserva = BuildJsonFail(em, logs)
End Function

