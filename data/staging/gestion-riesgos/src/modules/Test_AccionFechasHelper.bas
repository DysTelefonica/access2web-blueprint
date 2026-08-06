Attribute VB_Name = "Test_AccionFechasHelper"
Option Compare Database
Option Explicit

' ============================================================
' Test_AccionFechasHelper — TDD atoms for modAccionFechasHelper
'
' Helper: ValidarFechasAccion
'   Signature: Public Function ValidarFechasAccion( _
'                 ByVal p_FechaInicio As String
'                 ByVal p_FechaFinPrevista As String
'                 ByVal p_FechaBaseProyecto As String
'                 Optional ByRef p_Error As String) As Boolean
'
' Issue:   #90 — Punto 04 parte 2 — Constraints cronológicas
'          de fecha prevista e inicio (PMAccion / PCAccion).
'
' SDD:     e2e-form-by-form-2026-06-22 - Bloque 4 - REQ-CAL-04 (parte 2)
'
' Architecture:
'   * Helper puro (no DAO). La regla es función pura de
'     (p_FechaInicio, p_FechaFinPrevista, p_FechaBaseProyecto).
'     El cache-first se hace en el form layer (m_FechaBaseProyecto).
'   * Cache-first garantizado por el form: la query SQL a
'     TbProyectosEdiciones para resolver MIN(FechaEdicion) por
'     IDProyecto se hace UNA vez por form load. El helper nunca
'     consulta la BD.
'   * Tests verifican el contrato funcional (constraint (a), (b),
'     igualdad, sin FechaInicio) sin tocar BD: ningún fixture
'     ni override de db.
'
' Fixture IDs: NO APLICA — el helper es puro.
'
' Skill: access-vba-tdd v2.6.1
' ============================================================

' --- JSON helpers wrapper (project convention) ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' --- Helpers para formatear fechas como dd/mm/yyyy (formato VB Access) ---
Private Function FormatDate(ByVal p_Date As Date) As String
    FormatDate = Format$(p_Date, "dd/mm/yyyy")
End Function

' ----------------------------------------------------------------
' ATOM 1 — Sad: Constraint (a) violada.
'   FechaFinPrevista < FechaInicio → rechazar con mensaje claro.
' GIVEN FechaInicio=2026-08-15, FechaFinPrevista=2026-08-10
'       (FechaFinPrevista anterior a FechaInicio)
'       FechaBaseProyecto=2026-01-01 (no interviene en constraint a)
' WHEN  ValidarFechasAccion
' THEN  False, p_Error contiene "prevista" e "inicio" (mensaje en español).
'
' RED → GREEN con la regla (a): FechaFinPrevista >= FechaInicio.
' ----------------------------------------------------------------
Public Function Test_AccionFechasHelper_Sad_ConstraintA_FechaPrevistaAnteriorAInicio() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: FechaInicio=2026-08-15, FechaFinPrevista=2026-08-10, FechaBaseProyecto=2026-01-01"
    logs(1) = "2. Act: ValidarFechasAccion (constraint (a) violada)"
    logs(2) = "3. Assert: resultado = False"
    logs(3) = "4. Assert: p_Error contiene ""prevista"" y ""inicio"""

    Dim m_FechaInicio As String
    Dim m_FechaFinPrevista As String
    Dim m_FechaBaseProyecto As String
    Dim m_Resultado As Boolean
    Dim m_Error As String

    m_FechaInicio = "15/08/2026"
    m_FechaFinPrevista = "10/08/2026"
    m_FechaBaseProyecto = "01/01/2026"

    m_Resultado = ValidarFechasAccion( _
        m_FechaInicio, m_FechaFinPrevista, m_FechaBaseProyecto, m_Error)

    If m_Resultado <> False Then
        logs(3) = "3. Assert FAIL: esperaba False, obtuve " & CStr(m_Resultado) & _
            " (error=" & m_Error & ")"
        Test_AccionFechasHelper_Sad_ConstraintA_FechaPrevistaAnteriorAInicio = _
            BuildFail("Constraint (a) violada: FechaFinPrevista anterior a FechaInicio debe rechazarse", logs)
        Exit Function
    End If

    If Len(m_Error) = 0 Then
        logs(3) = "4. Assert FAIL: p_Error está vacío, esperaba mensaje claro en español"
        Test_AccionFechasHelper_Sad_ConstraintA_FechaPrevistaAnteriorAInicio = _
            BuildFail("p_Error debe estar poblado cuando el resultado es False", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "prevista", vbTextCompare) = 0 Then
        logs(3) = "4. Assert FAIL: p_Error no contiene ""prevista"": " & m_Error
        Test_AccionFechasHelper_Sad_ConstraintA_FechaPrevistaAnteriorAInicio = _
            BuildFail("p_Error debe mencionar la fecha prevista", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "inicio", vbTextCompare) = 0 Then
        logs(3) = "4. Assert FAIL: p_Error no contiene ""inicio"": " & m_Error
        Test_AccionFechasHelper_Sad_ConstraintA_FechaPrevistaAnteriorAInicio = _
            BuildFail("p_Error debe mencionar la fecha de inicio", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: False con mensaje claro en español"
    Test_AccionFechasHelper_Sad_ConstraintA_FechaPrevistaAnteriorAInicio = _
        BuildOk(CStr(m_Resultado), logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_AccionFechasHelper_Sad_ConstraintA_FechaPrevistaAnteriorAInicio = _
        BuildFail("ValidarFechasAccion raised: " & Err.description, logs)
End Function

' ----------------------------------------------------------------
' ATOM 2 — Edge: Constraint (a) satisfecha por igualdad.
'   FechaFinPrevista == FechaInicio → aceptar.
' GIVEN FechaInicio=FechaFinPrevista=2026-08-15
' WHEN  ValidarFechasAccion
' THEN  True, p_Error vacío.
' ----------------------------------------------------------------
Public Function Test_AccionFechasHelper_Edge_ConstraintA_IgualdadAceptada() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: FechaInicio=FechaFinPrevista=2026-08-15 (constraint (a) igualdad)"
    logs(1) = "2. Act: ValidarFechasAccion"
    logs(2) = "3. Assert: resultado = True"
    logs(3) = "4. Assert: p_Error vacío"

    Dim m_FechaInicio As String
    Dim m_FechaFinPrevista As String
    Dim m_FechaBaseProyecto As String
    Dim m_Resultado As Boolean
    Dim m_Error As String

    m_FechaInicio = "15/08/2026"
    m_FechaFinPrevista = "15/08/2026"
    m_FechaBaseProyecto = "01/01/2026"

    m_Resultado = ValidarFechasAccion( _
        m_FechaInicio, m_FechaFinPrevista, m_FechaBaseProyecto, m_Error)

    If m_Resultado <> True Then
        logs(3) = "3. Assert FAIL: esperaba True, obtuve " & CStr(m_Resultado) & _
            " (error=" & m_Error & ")"
        Test_AccionFechasHelper_Edge_ConstraintA_IgualdadAceptada = _
            BuildFail("Constraint (a) igualdad debe aceptarse", logs)
        Exit Function
    End If

    If Len(m_Error) <> 0 Then
        logs(3) = "4. Assert FAIL: p_Error poblado: " & m_Error
        Test_AccionFechasHelper_Edge_ConstraintA_IgualdadAceptada = _
            BuildFail("p_Error debe estar vacío cuando el resultado es True", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: True sin error (igualdad aceptada per spec)"
    Test_AccionFechasHelper_Edge_ConstraintA_IgualdadAceptada = _
        BuildOk(CStr(m_Resultado), logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_AccionFechasHelper_Edge_ConstraintA_IgualdadAceptada = _
        BuildFail("ValidarFechasAccion raised: " & Err.description, logs)
End Function

' ----------------------------------------------------------------
' ATOM 3 — Happy: Constraint (a) satisfecha por posterioridad.
'   FechaFinPrevista > FechaInicio → aceptar.
' GIVEN FechaInicio=2026-08-15, FechaFinPrevista=2026-09-30
' WHEN  ValidarFechasAccion
' THEN  True, p_Error vacío.
' ----------------------------------------------------------------
Public Function Test_AccionFechasHelper_Happy_ConstraintA_PosterioridadAceptada() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: FechaInicio=2026-08-15, FechaFinPrevista=2026-09-30 (constraint (a) posterior)"
    logs(1) = "2. Act: ValidarFechasAccion"
    logs(2) = "3. Assert: resultado = True"
    logs(3) = "4. Assert: p_Error vacío"

    Dim m_FechaInicio As String
    Dim m_FechaFinPrevista As String
    Dim m_FechaBaseProyecto As String
    Dim m_Resultado As Boolean
    Dim m_Error As String

    m_FechaInicio = "15/08/2026"
    m_FechaFinPrevista = "30/09/2026"
    m_FechaBaseProyecto = "01/01/2026"

    m_Resultado = ValidarFechasAccion( _
        m_FechaInicio, m_FechaFinPrevista, m_FechaBaseProyecto, m_Error)

    If m_Resultado <> True Then
        logs(3) = "3. Assert FAIL: esperaba True, obtuve " & CStr(m_Resultado) & _
            " (error=" & m_Error & ")"
        Test_AccionFechasHelper_Happy_ConstraintA_PosterioridadAceptada = _
            BuildFail("Constraint (a) posterior debe aceptarse", logs)
        Exit Function
    End If

    If Len(m_Error) <> 0 Then
        logs(3) = "4. Assert FAIL: p_Error poblado: " & m_Error
        Test_AccionFechasHelper_Happy_ConstraintA_PosterioridadAceptada = _
            BuildFail("p_Error debe estar vacío cuando el resultado es True", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: True sin error (posterioridad aceptada)"
    Test_AccionFechasHelper_Happy_ConstraintA_PosterioridadAceptada = _
        BuildOk(CStr(m_Resultado), logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_AccionFechasHelper_Happy_ConstraintA_PosterioridadAceptada = _
        BuildFail("ValidarFechasAccion raised: " & Err.description, logs)
End Function

' ----------------------------------------------------------------
' ATOM 4 — Sad: Constraint (b) violada.
'   FechaInicio < FechaBaseProyecto → rechazar con mensaje claro.
' GIVEN FechaBaseProyecto=2026-06-01, FechaInicio=2026-05-15 (anterior)
'       FechaFinPrevista=2026-07-15 (no interviene en constraint b)
' WHEN  ValidarFechasAccion
' THEN  False, p_Error contiene "inicio" y la base del proyecto.
' ----------------------------------------------------------------
Public Function Test_AccionFechasHelper_Sad_ConstraintB_FechaInicioAnteriorABaseProyecto() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: FechaBaseProyecto=2026-06-01, FechaInicio=2026-05-15 (anterior)"
    logs(1) = "2. Act: ValidarFechasAccion (constraint (b) violada)"
    logs(2) = "3. Assert: resultado = False"
    logs(3) = "4. Assert: p_Error contiene ""inicio"" y ""proyecto"" (o ""primera edición"")"

    Dim m_FechaInicio As String
    Dim m_FechaFinPrevista As String
    Dim m_FechaBaseProyecto As String
    Dim m_Resultado As Boolean
    Dim m_Error As String

    m_FechaInicio = "15/05/2026"
    m_FechaFinPrevista = "15/07/2026"
    m_FechaBaseProyecto = "01/06/2026"

    m_Resultado = ValidarFechasAccion( _
        m_FechaInicio, m_FechaFinPrevista, m_FechaBaseProyecto, m_Error)

    If m_Resultado <> False Then
        logs(3) = "3. Assert FAIL: esperaba False, obtuve " & CStr(m_Resultado) & _
            " (error=" & m_Error & ")"
        Test_AccionFechasHelper_Sad_ConstraintB_FechaInicioAnteriorABaseProyecto = _
            BuildFail("Constraint (b) violada: FechaInicio anterior a FechaBaseProyecto debe rechazarse", logs)
        Exit Function
    End If

    If Len(m_Error) = 0 Then
        logs(3) = "4. Assert FAIL: p_Error está vacío, esperaba mensaje claro en español"
        Test_AccionFechasHelper_Sad_ConstraintB_FechaInicioAnteriorABaseProyecto = _
            BuildFail("p_Error debe estar poblado cuando el resultado es False", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "inicio", vbTextCompare) = 0 Then
        logs(3) = "4. Assert FAIL: p_Error no contiene ""inicio"": " & m_Error
        Test_AccionFechasHelper_Sad_ConstraintB_FechaInicioAnteriorABaseProyecto = _
            BuildFail("p_Error debe mencionar la fecha de inicio", logs)
        Exit Function
    End If

    If (InStr(1, m_Error, "proyecto", vbTextCompare) = 0) And _
       (InStr(1, m_Error, "primera", vbTextCompare) = 0) Then
        logs(3) = "4. Assert FAIL: p_Error no contiene ""proyecto"" ni ""primera"": " & m_Error
        Test_AccionFechasHelper_Sad_ConstraintB_FechaInicioAnteriorABaseProyecto = _
            BuildFail("p_Error debe mencionar el proyecto o la primera edición", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: False con mensaje claro en español"
    Test_AccionFechasHelper_Sad_ConstraintB_FechaInicioAnteriorABaseProyecto = _
        BuildOk(CStr(m_Resultado), logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_AccionFechasHelper_Sad_ConstraintB_FechaInicioAnteriorABaseProyecto = _
        BuildFail("ValidarFechasAccion raised: " & Err.description, logs)
End Function

' ----------------------------------------------------------------
' ATOM 5 — Happy: Constraint (b) satisfecha.
'   FechaInicio >= FechaBaseProyecto → aceptar.
' GIVEN FechaBaseProyecto=2026-06-01, FechaInicio=2026-06-01 (igualdad)
' WHEN  ValidarFechasAccion
' THEN  True, p_Error vacío.
' ----------------------------------------------------------------
Public Function Test_AccionFechasHelper_Happy_ConstraintB_SatisfechaIgualdad() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: FechaBaseProyecto=FechaInicio=2026-06-01 (constraint (b) igualdad)"
    logs(1) = "2. Act: ValidarFechasAccion"
    logs(2) = "3. Assert: resultado = True"
    logs(3) = "4. Assert: p_Error vacío"

    Dim m_FechaInicio As String
    Dim m_FechaFinPrevista As String
    Dim m_FechaBaseProyecto As String
    Dim m_Resultado As Boolean
    Dim m_Error As String

    m_FechaInicio = "01/06/2026"
    m_FechaFinPrevista = "30/06/2026"
    m_FechaBaseProyecto = "01/06/2026"

    m_Resultado = ValidarFechasAccion( _
        m_FechaInicio, m_FechaFinPrevista, m_FechaBaseProyecto, m_Error)

    If m_Resultado <> True Then
        logs(3) = "3. Assert FAIL: esperaba True, obtuve " & CStr(m_Resultado) & _
            " (error=" & m_Error & ")"
        Test_AccionFechasHelper_Happy_ConstraintB_SatisfechaIgualdad = _
            BuildFail("Constraint (b) igualdad debe aceptarse", logs)
        Exit Function
    End If

    If Len(m_Error) <> 0 Then
        logs(3) = "4. Assert FAIL: p_Error poblado: " & m_Error
        Test_AccionFechasHelper_Happy_ConstraintB_SatisfechaIgualdad = _
            BuildFail("p_Error debe estar vacío cuando el resultado es True", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: True sin error (constraint (b) igualdad aceptada)"
    Test_AccionFechasHelper_Happy_ConstraintB_SatisfechaIgualdad = _
        BuildOk(CStr(m_Resultado), logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_AccionFechasHelper_Happy_ConstraintB_SatisfechaIgualdad = _
        BuildFail("ValidarFechasAccion raised: " & Err.description, logs)
End Function

' ----------------------------------------------------------------
' ATOM 6 — Edge: Sin FechaInicio → constraint (a) no aplica.
' GIVEN FechaInicio="" (vacía), FechaFinPrevista=2026-08-15
'       FechaBaseProyecto=2026-06-01
' WHEN  ValidarFechasAccion
' THEN  True, p_Error vacío.
'
' Spec del issue #90 escenario 6: "Sin FechaInicio: constraint (a)
' no aplica (devolver True)". El form aplica la regla ""requerido""
' por separado; nuestro helper NO rechaza fechas vacías.
' ----------------------------------------------------------------
Public Function Test_AccionFechasHelper_Edge_SinFechaInicio_ConstraintANoAplica() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: FechaInicio="""" (vacía), FechaFinPrevista=2026-08-15, FechaBaseProyecto=2026-06-01"
    logs(1) = "2. Act: ValidarFechasAccion"
    logs(2) = "3. Assert: resultado = True (constraint (a) no aplica)"
    logs(3) = "4. Assert: p_Error vacío"

    Dim m_FechaInicio As String
    Dim m_FechaFinPrevista As String
    Dim m_FechaBaseProyecto As String
    Dim m_Resultado As Boolean
    Dim m_Error As String

    m_FechaInicio = ""
    m_FechaFinPrevista = "15/08/2026"
    m_FechaBaseProyecto = "01/06/2026"

    m_Resultado = ValidarFechasAccion( _
        m_FechaInicio, m_FechaFinPrevista, m_FechaBaseProyecto, m_Error)

    If m_Resultado <> True Then
        logs(3) = "3. Assert FAIL: esperaba True, obtuve " & CStr(m_Resultado) & _
            " (error=" & m_Error & ")"
        Test_AccionFechasHelper_Edge_SinFechaInicio_ConstraintANoAplica = _
            BuildFail("Sin FechaInicio, constraint (a) no debe aplicar", logs)
        Exit Function
    End If

    If Len(m_Error) <> 0 Then
        logs(3) = "4. Assert FAIL: p_Error poblado: " & m_Error
        Test_AccionFechasHelper_Edge_SinFechaInicio_ConstraintANoAplica = _
            BuildFail("p_Error debe estar vacío cuando el resultado es True", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: True sin error (sin FechaInicio no se rechaza)"
    Test_AccionFechasHelper_Edge_SinFechaInicio_ConstraintANoAplica = _
        BuildOk(CStr(m_Resultado), logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_AccionFechasHelper_Edge_SinFechaInicio_ConstraintANoAplica = _
        BuildFail("ValidarFechasAccion raised: " & Err.description, logs)
End Function

' ----------------------------------------------------------------
' ATOM 7 — Adversarial: FechaBaseProyecto vacía.
' Si el form no resolvió m_FechaBaseProyecto (form layer no cacheó),
' el helper no debe crashear y debe devolver True para constraint (b)
' (no se puede validar sin el dato). Constraint (a) sigue aplicando
' normalmente si FechaInicio está presente.
'
' GIVEN FechaBaseProyecto="", FechaInicio=2026-08-15, FechaFinPrevista=2026-09-15
' WHEN  ValidarFechasAccion
' THEN  True, p_Error vacío (skip de constraint b por falta de dato).
' ----------------------------------------------------------------
Public Function Test_AccionFechasHelper_Adversarial_FechaBaseProyectoVacia_NoCrashea() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: FechaBaseProyecto="""" (vacía), FechaInicio=2026-08-15, FechaFinPrevista=2026-09-15"
    logs(1) = "2. Act: ValidarFechasAccion (form no resolvió cache)"
    logs(2) = "3. Assert: resultado = True (no podemos validar constraint b sin dato)"
    logs(3) = "4. Assert: p_Error vacío (no es error nuestro)"

    Dim m_FechaInicio As String
    Dim m_FechaFinPrevista As String
    Dim m_FechaBaseProyecto As String
    Dim m_Resultado As Boolean
    Dim m_Error As String

    m_FechaInicio = "15/08/2026"
    m_FechaFinPrevista = "15/09/2026"
    m_FechaBaseProyecto = ""

    m_Resultado = ValidarFechasAccion( _
        m_FechaInicio, m_FechaFinPrevista, m_FechaBaseProyecto, m_Error)

    If m_Resultado <> True Then
        logs(3) = "3. Assert FAIL: esperaba True, obtuve " & CStr(m_Resultado) & _
            " (error=" & m_Error & ")"
        Test_AccionFechasHelper_Adversarial_FechaBaseProyectoVacia_NoCrashea = _
            BuildFail("Sin FechaBaseProyecto, constraint (b) no aplica (no rechazar)", logs)
        Exit Function
    End If

    If Len(m_Error) <> 0 Then
        logs(3) = "4. Assert FAIL: p_Error poblado: " & m_Error
        Test_AccionFechasHelper_Adversarial_FechaBaseProyectoVacia_NoCrashea = _
            BuildFail("p_Error debe estar vacío cuando no podemos validar (b)", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: True sin error (skip defensivo de (b) sin cache)"
    Test_AccionFechasHelper_Adversarial_FechaBaseProyectoVacia_NoCrashea = _
        BuildOk(CStr(m_Resultado), logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_AccionFechasHelper_Adversarial_FechaBaseProyectoVacia_NoCrashea = _
        BuildFail("ValidarFechasAccion raised: " & Err.description, logs)
End Function

' ----------------------------------------------------------------
' ATOM 8 — Adversarial: FechaFinPrevista inválida con FechaInicio válida.
' Si el form envía una FechaFinPrevista que no es fecha (string no
' parseable), el helper NO debe crashear y debe devolver True para
' constraint (a). El form layer es responsable de validar tipos.
'
' GIVEN FechaInicio=2026-08-15, FechaFinPrevista="no es fecha"
' WHEN  ValidarFechasAccion
' THEN  True, p_Error vacío (skip defensivo).
' ----------------------------------------------------------------
Public Function Test_AccionFechasHelper_Adversarial_FechaFinPrevistaInvalida_NoCrashea() As String
    On Error GoTo EH

    Dim logs(0 To 4) As String
    logs(0) = "1. Arrange: FechaInicio=2026-08-15, FechaFinPrevista=""no es fecha"""
    logs(1) = "2. Act: ValidarFechasAccion (string no parseable)"
    logs(2) = "3. Assert: resultado = True (no es nuestro problema)"
    logs(3) = "4. Assert: p_Error vacío"

    Dim m_FechaInicio As String
    Dim m_FechaFinPrevista As String
    Dim m_FechaBaseProyecto As String
    Dim m_Resultado As Boolean
    Dim m_Error As String

    m_FechaInicio = "15/08/2026"
    m_FechaFinPrevista = "no es fecha"
    m_FechaBaseProyecto = "01/06/2026"

    m_Resultado = ValidarFechasAccion( _
        m_FechaInicio, m_FechaFinPrevista, m_FechaBaseProyecto, m_Error)

    If m_Resultado <> True Then
        logs(3) = "3. Assert FAIL: esperaba True, obtuve " & CStr(m_Resultado) & _
            " (error=" & m_Error & ")"
        Test_AccionFechasHelper_Adversarial_FechaFinPrevistaInvalida_NoCrashea = _
            BuildFail("Fecha inválida en form layer es problema del form, no del helper", logs)
        Exit Function
    End If

    If Len(m_Error) <> 0 Then
        logs(3) = "4. Assert FAIL: p_Error poblado: " & m_Error
        Test_AccionFechasHelper_Adversarial_FechaFinPrevistaInvalida_NoCrashea = _
            BuildFail("p_Error debe estar vacío para fecha no parseable", logs)
        Exit Function
    End If

    logs(3) = "4. Assert PASS: True sin error (skip defensivo de tipo)"
    Test_AccionFechasHelper_Adversarial_FechaFinPrevistaInvalida_NoCrashea = _
        BuildOk(CStr(m_Resultado), logs)
    Exit Function

EH:
    logs(2) = "2. Act: raised " & Err.Number & " - " & Err.description
    Test_AccionFechasHelper_Adversarial_FechaFinPrevistaInvalida_NoCrashea = _
        BuildFail("ValidarFechasAccion raised: " & Err.description, logs)
End Function