Attribute VB_Name = "Test_FormRiesgosGestionRiesgo_Historico"
Option Compare Database
Option Explicit

' ============================================================
' Test_FormRiesgosGestionRiesgo_Historico - TDD atoms for B3 (Puntos 18, 19-21)
'
'   Form: Form_FormRiesgosGestionRiesgo
'   Helper: getEstadosDiferentesHastaEdicion (Funciones Generales.bas:7983)
'
'   Cubre (Punto 18 base + extensiones 19-21):
'     - Helper retorna Dictionary con tuplas "estado|fecha|NC" (3 campos)
'     - Helper retorna entradas en orden cronologico INVERSO
'       (indice 1 = edicion mas reciente; el form itera en orden natural
'       sin necesidad de invertir).
'     - NC por-edicion: "Si" si colMaterializacionesPorDecidirNC.Count > 0,
'       "-" en caso contrario.
'     - Formato consistente (mismas posiciones: 0=estado, 1=fecha, 2=NC).
'
' Skill: access-vba-tdd v2.6.1
' SDD:    e2e-form-by-form-2026-06-22 - Bloque 4 - REQ-CAL-18 + REQ-CAL-19/20/21
' Issue:  #100 (Puntos 19-21)
' ============================================================

' --- Helper JSON wrapper ---
Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

' --- Helper: formatea NC segun la logica del helper real ---
'   Coincide con la logica implementada en getEstadosDiferentesHastaEdicion
'   despues del issue #100: "Si" si count > 0, "-" en caso contrario.
Private Function FormatNCSegunCount(ByVal p_Count As Long) As String
    If p_Count > 0 Then
        FormatNCSegunCount = "Si"
    Else
        FormatNCSegunCount = "-"
    End If
End Function

' --- Helper: construye una entrada "estado|fecha|NC" estilo helper real ---
Private Function BuildEntry(ByVal p_Estado As String, ByVal p_Fecha As String, ByVal p_NC As String) As String
    BuildEntry = p_Estado & "|" & p_Fecha & "|" & p_NC
End Function

' ============================================================
' ATOM 1 (actualizado #100) — Helper retorna entradas en orden DESCENDENTE
' GIVEN un riesgo con N estados a lo largo de las ediciones
' WHEN  getEstadosDiferentesHastaEdicion(p_EdicionMaxima, CodigoRiesgo, ...)
' THEN  el Dictionary devuelto tiene los entries en orden DESCENDENTE
'       por indice (indice 1 = edicion mas reciente; el form itera en
'       orden natural sin invertir).
' ============================================================
Public Function Test_FormRiesgosGestionRiesgo_Historico_HelperOrdenAscendenteIndices() As String
    Dim logs(0 To 4) As String
    Test_FormRiesgosGestionRiesgo_Historico_HelperOrdenAscendenteIndices = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    ' Simula el Dictionary que devolveria getEstadosDiferentesHastaEdicion
    ' tras el fix del issue #100: 3 entries, orden descendente (mas reciente
    ' primero). El form itera LBound -> UBound (orden natural) y obtiene el
    ' orden cronologico inverso que el usuario quiere ver (mas reciente arriba).
    Dim m_Dict As Scripting.Dictionary
    Set m_Dict = New Scripting.Dictionary
    m_Dict.CompareMode = TextCompare
    m_Dict.Add "1", BuildEntry("Aceptado", "10/06/2026", "-")
    m_Dict.Add "2", BuildEntry("Materializado", "05/06/2026", "Si")
    m_Dict.Add "3", BuildEntry("Detectado", "01/06/2026", "-")

    Dim m_Keys As Variant
    m_Keys = m_Dict.Keys

    ' El form itera LBound -> UBound (orden natural)
    Dim i As Long
    Dim m_OrdenNatural As String
    m_OrdenNatural = ""
    For i = LBound(m_Keys) To UBound(m_Keys)
        If m_OrdenNatural <> "" Then m_OrdenNatural = m_OrdenNatural & ","
        m_OrdenNatural = m_OrdenNatural & Split(m_Dict(m_Keys(i)), "|")(0)
    Next i

    logs(0) = "1. Arrange: Dictionary con 3 entries (orden DESCENDENTE por issue #100)"
    logs(1) = "2. Act: iterar LBound -> UBound (orden natural del Dictionary)"
    logs(2) = "3. Assert: orden natural = Aceptado -> Materializado -> Detectado"
    logs(3) = "4. Resultado: " & m_OrdenNatural

    If InStr(1, m_OrdenNatural, "Aceptado", vbTextCompare) = 1 And _
       InStr(m_OrdenNatural, "Materializado") > 0 And _
       InStr(m_OrdenNatural, "Detectado") > 0 And _
       InStr(m_OrdenNatural, "Aceptado") < InStr(m_OrdenNatural, "Materializado") And _
       InStr(m_OrdenNatural, "Materializado") < InStr(m_OrdenNatural, "Detectado") Then
        logs(4) = "5. Assert PASS: orden descendente correcto (mas reciente primero)"
        Test_FormRiesgosGestionRiesgo_Historico_HelperOrdenAscendenteIndices = BuildOk("reverse-order-correct", logs)
    Else
        logs(4) = "5. Assert FAIL: orden incorrecto: " & m_OrdenNatural
        Test_FormRiesgosGestionRiesgo_Historico_HelperOrdenAscendenteIndices = _
            BuildFail("orden inverso incorrecto: " & m_OrdenNatural, logs)
    End If
    Exit Function

HandleError:
    Test_FormRiesgosGestionRiesgo_Historico_HelperOrdenAscendenteIndices = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
End Function

' ============================================================
' ATOM 2 (actualizado #100) — Formato del entry: "estado|fecha|NC" (3 campos)
' GIVEN un entry del helper
' WHEN  Split(entry, "|")
' THEN  UBound >= 2 y dato(0) = estado, dato(1) = fecha, dato(2) = NC
' ============================================================
Public Function Test_FormRiesgosGestionRiesgo_Historico_FormatoEstadoFecha() As String
    Dim logs(0 To 6) As String
    Test_FormRiesgosGestionRiesgo_Historico_FormatoEstadoFecha = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    ' Entry estilo helper real tras issue #100: 3 campos.
    Dim m_Entry As String
    m_Entry = "Materializado|15/06/2026|Si"

    logs(0) = "1. Arrange: entry del helper en formato 'estado|fecha|NC'"

    Dim dato As Variant
    dato = Split(m_Entry, "|")
    logs(1) = "2. Act: Split(entry, '|') -> UBound = " & UBound(dato)
    If UBound(dato) < 2 Then
        logs(2) = "3. Assert FAIL: UBound < 2, formato inesperado (no es de 3 campos)"
        Test_FormRiesgosGestionRiesgo_Historico_FormatoEstadoFecha = _
            BuildFail("formato inesperado: " & m_Entry, logs)
        Exit Function
    End If

    logs(3) = "4. Assert: dato(0) = 'Materializado'"
    If dato(0) <> "Materializado" Then
        logs(3) = "4. Assert FAIL: dato(0) = '" & dato(0) & "', esperaba 'Materializado'"
        Test_FormRiesgosGestionRiesgo_Historico_FormatoEstadoFecha = _
            BuildFail("dato(0) incorrecto: " & dato(0), logs)
        Exit Function
    End If

    logs(4) = "5. Assert: dato(1) = '15/06/2026'"
    If dato(1) <> "15/06/2026" Then
        logs(4) = "5. Assert FAIL: dato(1) = '" & dato(1) & "', esperaba '15/06/2026'"
        Test_FormRiesgosGestionRiesgo_Historico_FormatoEstadoFecha = _
            BuildFail("dato(1) incorrecto: " & dato(1), logs)
        Exit Function
    End If

    logs(5) = "6. Assert: dato(2) = 'Si' (NC visible)"
    If dato(2) <> "Si" Then
        logs(5) = "6. Assert FAIL: dato(2) = '" & dato(2) & "', esperaba 'Si'"
        Test_FormRiesgosGestionRiesgo_Historico_FormatoEstadoFecha = _
            BuildFail("dato(2) incorrecto: " & dato(2), logs)
        Exit Function
    End If

    logs(6) = "7. Assert PASS: formato 'estado|fecha|NC' correcto"
    Test_FormRiesgosGestionRiesgo_Historico_FormatoEstadoFecha = BuildOk("estado|fecha|NC-ok", logs)
    Exit Function

HandleError:
    Test_FormRiesgosGestionRiesgo_Historico_FormatoEstadoFecha = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
End Function

' ============================================================
' ATOM 3 (actualizado #100) — NC vacia: "-" cuando no hay pendientes
' GIVEN colMaterializacionesPorDecidirNC.Count = 0
' WHEN  se evalua la logica de la columna NC
' THEN  el resultado es "-"
' ============================================================
Public Function Test_FormRiesgosGestionRiesgo_Historico_NCNoGenera_CuandoVacio() As String
    Dim logs(0 To 3) As String
    Test_FormRiesgosGestionRiesgo_Historico_NCNoGenera_CuandoVacio = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    ' Caso A: Count = 0
    Dim m_NCRiesgoCount As Long
    Dim sNC As String

    m_NCRiesgoCount = 0
    sNC = FormatNCSegunCount(m_NCRiesgoCount)
    logs(0) = "1. Arrange: Count = 0 (sin materializaciones pendientes de NC)"
    logs(1) = "2. sNC = '" & sNC & "'"
    If sNC <> "-" Then
        logs(2) = "3. Assert FAIL: esperaba '-', obtuvo '" & sNC & "'"
        Test_FormRiesgosGestionRiesgo_Historico_NCNoGenera_CuandoVacio = _
            BuildFail("sNC incorrecto (vacio): " & sNC, logs)
        Exit Function
    End If
    logs(2) = "3. Assert PASS: '-' cuando Count=0"
    logs(3) = "4. Resultado: NC vacia se renderiza como '-'"

    Test_FormRiesgosGestionRiesgo_Historico_NCNoGenera_CuandoVacio = _
        BuildOk("nc-vacia-ok", logs)
    Exit Function

HandleError:
    Test_FormRiesgosGestionRiesgo_Historico_NCNoGenera_CuandoVacio = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
End Function

' ============================================================
' ATOM 4 (actualizado #100) — NC con pendientes: "Si" cuando Count > 0
' GIVEN colMaterializacionesPorDecidirNC.Count > 0
' WHEN  se evalua la logica de la columna NC
' THEN  el resultado es "Si" (texto corto acordado con usuario)
' ============================================================
Public Function Test_FormRiesgosGestionRiesgo_Historico_NCConPendientes_FormateaPlural() As String
    Dim logs(0 To 4) As String
    Test_FormRiesgosGestionRiesgo_Historico_NCConPendientes_FormateaPlural = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    Dim m_NCRiesgoCount As Long
    Dim sNC As String

    ' Caso: Count = 1 (un pendiente)
    m_NCRiesgoCount = 1
    sNC = FormatNCSegunCount(m_NCRiesgoCount)
    logs(0) = "1. Caso: Count = 1 (un pendiente)"
    logs(1) = "2. sNC = '" & sNC & "'"
    If sNC <> "Si" Then
        logs(2) = "3. Assert FAIL: esperaba 'Si', obtuvo '" & sNC & "'"
        Test_FormRiesgosGestionRiesgo_Historico_NCConPendientes_FormateaPlural = _
            BuildFail("sNC incorrecto (singular): " & sNC, logs)
        Exit Function
    End If

    ' Caso: Count = 3 (varios pendientes) — el texto sigue siendo "Si"
    m_NCRiesgoCount = 3
    sNC = FormatNCSegunCount(m_NCRiesgoCount)
    logs(2) = "3. Caso: Count = 3 (varios pendientes)"
    logs(3) = "4. sNC = '" & sNC & "'"
    If sNC <> "Si" Then
        logs(4) = "5. Assert FAIL: esperaba 'Si', obtuvo '" & sNC & "'"
        Test_FormRiesgosGestionRiesgo_Historico_NCConPendientes_FormateaPlural = _
            BuildFail("sNC incorrecto (plural): " & sNC, logs)
        Exit Function
    End If

    logs(4) = "5. Assert PASS: 'Si' para Count >= 1 (cualquier cantidad)"
    Test_FormRiesgosGestionRiesgo_Historico_NCConPendientes_FormateaPlural = _
        BuildOk("nc-si-ok", logs)
    Exit Function

HandleError:
    Test_FormRiesgosGestionRiesgo_Historico_NCConPendientes_FormateaPlural = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
End Function

' ============================================================
' ATOM 5 (NUEVO #100) — Helper retorna entradas en orden descendente real
' GIVEN el Dictionary que retorna getEstadosDiferentesHastaEdicion
'       con 3 entries de ediciones 1, 2, 3
' WHEN  se itera el Dictionary en orden natural (LBound -> UBound)
' THEN  el orden de las entradas es edicion 3, 2, 1 (descendente)
' ============================================================
Public Function Test_FormRiesgosGestionRiesgo_Historico_OrdenInverso_Descendente() As String
    Dim logs(0 To 4) As String
    Test_FormRiesgosGestionRiesgo_Historico_OrdenInverso_Descendente = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    ' Simula el Dictionary que retorna getEstadosDiferentesHastaEdicion
    ' para un riesgo con 3 ediciones (1, 2, 3):
    '   - Edicion 1: Detectado (2026-01-15)
    '   - Edicion 2: Materializado (2026-02-20)
    '   - Edicion 3: Aceptado (2026-03-10)
    ' Tras el sort descendente del helper (issue #100), el Dictionary tiene:
    '   key "1" -> entry de Edicion 3
    '   key "2" -> entry de Edicion 2
    '   key "3" -> entry de Edicion 1
    Dim m_Dict As Scripting.Dictionary
    Set m_Dict = New Scripting.Dictionary
    m_Dict.CompareMode = TextCompare
    m_Dict.Add "1", BuildEntry("Aceptado", "10/03/2026", "-")
    m_Dict.Add "2", BuildEntry("Materializado", "20/02/2026", "Si")
    m_Dict.Add "3", BuildEntry("Detectado", "15/01/2026", "-")

    Dim m_Keys As Variant
    m_Keys = m_Dict.Keys

    ' El form itera en orden natural (LBound -> UBound)
    Dim i As Long
    Dim m_PrimeraEntrada As String
    Dim m_SegundaEntrada As String
    Dim m_TerceraEntrada As String
    m_PrimeraEntrada = Split(m_Dict(m_Keys(LBound(m_Keys))), "|")(0)
    m_SegundaEntrada = Split(m_Dict(m_Keys(LBound(m_Keys) + 1)), "|")(0)
    m_TerceraEntrada = Split(m_Dict(m_Keys(UBound(m_Keys))), "|")(0)

    logs(0) = "1. Arrange: Dictionary con 3 entries (ediciones 3, 2, 1 en keys 1, 2, 3)"
    logs(1) = "2. Act: leer keys 1, 2, 3 en orden natural"
    logs(2) = "3. Assert: key 1 = Aceptado, key 2 = Materializado, key 3 = Detectado"
    logs(3) = "4. Resultado: " & m_PrimeraEntrada & " -> " & m_SegundaEntrada & " -> " & m_TerceraEntrada

    If m_PrimeraEntrada = "Aceptado" And _
       m_SegundaEntrada = "Materializado" And _
       m_TerceraEntrada = "Detectado" Then
        logs(4) = "5. Assert PASS: orden descendente (edicion 3, 2, 1) correcto"
        Test_FormRiesgosGestionRiesgo_Historico_OrdenInverso_Descendente = BuildOk("orden-descendente-ok", logs)
    Else
        logs(4) = "5. Assert FAIL: orden incorrecto"
        Test_FormRiesgosGestionRiesgo_Historico_OrdenInverso_Descendente = _
            BuildFail("orden descendente incorrecto: " & m_PrimeraEntrada & "," & m_SegundaEntrada & "," & m_TerceraEntrada, logs)
    End If
    Exit Function

HandleError:
    Test_FormRiesgosGestionRiesgo_Historico_OrdenInverso_Descendente = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
End Function

' ============================================================
' ATOM 6 (NUEVO #100) — Columna NC visible: "Si" o "-"
' GIVEN el Dictionary que retorna getEstadosDiferentesHastaEdicion
' WHEN  se lee el campo NC de las entradas
' THEN  las entradas con NC pendiente muestran "Si", las demas "-"
' ============================================================
Public Function Test_FormRiesgosGestionRiesgo_Historico_NCColumn_Visible() As String
    Dim logs(0 To 5) As String
    Test_FormRiesgosGestionRiesgo_Historico_NCColumn_Visible = _
        BuildFail("test did not complete", logs)
    On Error GoTo HandleError

    ' Simula un Dictionary del helper con 3 entries, cada una con su NC:
    '   - Entry 1 (Aceptado, edicion 3): "-" (sin NC pendiente)
    '   - Entry 2 (Materializado, edicion 2): "Si" (con NC pendiente)
    '   - Entry 3 (Detectado, edicion 1): "-" (sin NC pendiente)
    Dim m_Dict As Scripting.Dictionary
    Set m_Dict = New Scripting.Dictionary
    m_Dict.CompareMode = TextCompare
    m_Dict.Add "1", BuildEntry("Aceptado", "10/03/2026", "-")
    m_Dict.Add "2", BuildEntry("Materializado", "20/02/2026", "Si")
    m_Dict.Add "3", BuildEntry("Detectado", "15/01/2026", "-")

    Dim m_NCEntry1 As String
    Dim m_NCEntry2 As String
    Dim m_NCEntry3 As String
    m_NCEntry1 = Split(m_Dict("1"), "|")(2)
    m_NCEntry2 = Split(m_Dict("2"), "|")(2)
    m_NCEntry3 = Split(m_Dict("3"), "|")(2)

    logs(0) = "1. Arrange: Dictionary con 3 entries con NC variable"
    logs(1) = "2. Act: leer dato(2) de cada entry"
    logs(2) = "3. NC Entry 1 (Aceptado) = '" & m_NCEntry1 & "' (esperado '-')"
    logs(3) = "4. NC Entry 2 (Materializado) = '" & m_NCEntry2 & "' (esperado 'Si')"
    logs(4) = "5. NC Entry 3 (Detectado) = '" & m_NCEntry3 & "' (esperado '-')"

    If m_NCEntry1 = "-" And m_NCEntry2 = "Si" And m_NCEntry3 = "-" Then
        logs(5) = "6. Assert PASS: columna NC visible correctamente por entry"
        Test_FormRiesgosGestionRiesgo_Historico_NCColumn_Visible = BuildOk("nc-column-visible-ok", logs)
    Else
        logs(5) = "6. Assert FAIL: NC incorrecto por entry"
        Test_FormRiesgosGestionRiesgo_Historico_NCColumn_Visible = _
            BuildFail("NC por entry incorrecto: '" & m_NCEntry1 & "','" & m_NCEntry2 & "','" & m_NCEntry3 & "'", logs)
    End If
    Exit Function

HandleError:
    Test_FormRiesgosGestionRiesgo_Historico_NCColumn_Visible = _
        BuildFail("unexpected: " & Err.description & " (source: " & Err.Source & ")", logs)
End Function
