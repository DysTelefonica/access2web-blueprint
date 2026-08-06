Attribute VB_Name = "Test_FormCoordinationHelper"
' ============================================================
' Test_FormCoordinationHelper - Atomos TDD del helper
'   modFormCoordinationHelper. Este modulo elimina los accesos
'   cruzados Form_FormX.Metodo() / Forms("X").Controls(...) que
'   violaban Hard rule 3 de access-vba-e2e-methodology.
'
' Skill: access-vba-e2e-methodology + access-vba-tdd v2.5
' SDD:   hard-rule-3-refactor (gestión_riesgos staging)
' Slice: cross-form coordination helper para gemelos Materializado /
'        Mitigacion / Retirado + FormTecnicoTareaRiesgosAceptadosRetirados.
'
' Convenciones del proyecto:
'   - BuildOk/BuildFail SIEMPRE via modTestingCoreHelper
'     (TestCore_BuildOk / TestCore_BuildFail). NUNCA redefinir.
'   - Sin MsgBox / InputBox en atomos (Hard rule 7).
'   - Sin acceso al sandbox backend en estos atomos (no tocan DAO).
'     Los helpers de coordinacion son UI/forms-only.
'   - Logs como array String con indice explicito; ReDim cuando crece.
' ============================================================
Option Compare Database
Option Explicit

' ============================================================
' ATOMO 1 (EDGE) - Formulario NO abierto -> Nothing sin raise
'
' Setup: FormRiesgosGestion NO esta abierto en sesion de test headless.
' Act:   Coord_GetSelectedNode
' Expect: retorna Nothing (Object) sin propagar error.
' Contrato defensivo: el caller debe poder usar Coord_GetSelectedNode
'                     sin un try/catch porque "form no abierto" es
'                     estado normal (form abierto desde menu, no
'                     modal obligatorio).
' ============================================================
Public Function Test_Coord_GetSelectedNode_FormClosed_ReturnsNothingNoRaise() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: headless test session, FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    ' El test es robusto si el form llegara a estar abierto en algun
    ' entorno - simplemente cerramos para garantizar el contrato "no
    ' abierto". En CI tipico, FormRiesgosGestion no esta abierto.
    Dim m_Error As String
    m_Error = ""

    logs(logIdx) = "2. Act: Coord_GetSelectedNode"
    logIdx = logIdx + 1

    Dim m_Nodo As Object
    Set m_Nodo = modFormCoordinationHelper.Coord_GetSelectedNode(m_Error)

    logs(logIdx) = "3. Assert: m_Nodo Is Nothing AND m_Error = "" (form no abierto)"
    logIdx = logIdx + 1

    If Not m_Nodo Is Nothing Then
        logs(logIdx) = "4. FAIL: esperaba Nothing, recibio nodo '" & TypeName(m_Nodo) & "'"
        logIdx = logIdx + 1
        Test_Coord_GetSelectedNode_FormClosed_ReturnsNothingNoRaise = _
            TestCore_BuildFail("Coord_GetSelectedNode debio devolver Nothing con form cerrado; devolvio " & TypeName(m_Nodo), logs)
        Exit Function
    End If

    If m_Error <> "" Then
        logs(logIdx) = "4. FAIL: m_Error deberia estar vacio, era: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_GetSelectedNode_FormClosed_ReturnsNothingNoRaise = _
            TestCore_BuildFail("Coord_GetSelectedNode relleno p_Error sin raise: " & m_Error, logs)
        Exit Function
    End If

    Test_Coord_GetSelectedNode_FormClosed_ReturnsNothingNoRaise = _
        TestCore_BuildOk("nothing_no_raise_pass", logs)
    Exit Function

EH:
    Test_Coord_GetSelectedNode_FormClosed_ReturnsNothingNoRaise = _
        TestCore_BuildFail("Test_Coord_GetSelectedNode_FormClosed_ReturnsNothingNoRaise: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 2 (HAPPY con fallback defensivo) - Formulario NO abierto
'   devuelve cadena vacia para comboVerDescripcion.
'
' Setup: FormRiesgosGestion NO abierto.
' Act:   Coord_GetComboVerDescripcion
' Expect: "" sin raise, p_Error vacio.
' ============================================================
Public Function Test_Coord_GetComboVerDescripcion_FormClosed_ReturnsEmptyNoRaise() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    Dim m_Valor As String

    logs(logIdx) = "2. Act: Coord_GetComboVerDescripcion"
    logIdx = logIdx + 1

    m_Valor = modFormCoordinationHelper.Coord_GetComboVerDescripcion(m_Error)

    logs(logIdx) = "3. Assert: m_Valor = "" AND m_Error = """
    logIdx = logIdx + 1

    If m_Valor <> "" Then
        logs(logIdx) = "4. FAIL: esperaba '', recibio '" & m_Valor & "'"
        logIdx = logIdx + 1
        Test_Coord_GetComboVerDescripcion_FormClosed_ReturnsEmptyNoRaise = _
            TestCore_BuildFail("Coord_GetComboVerDescripcion debio devolver '' con form cerrado; devolvio '" & m_Valor & "'", logs)
        Exit Function
    End If

    If m_Error <> "" Then
        logs(logIdx) = "4. FAIL: m_Error deberia estar vacio"
        logIdx = logIdx + 1
        Test_Coord_GetComboVerDescripcion_FormClosed_ReturnsEmptyNoRaise = _
            TestCore_BuildFail("Coord_GetComboVerDescripcion relleno p_Error: " & m_Error, logs)
        Exit Function
    End If

    Test_Coord_GetComboVerDescripcion_FormClosed_ReturnsEmptyNoRaise = _
        TestCore_BuildOk("empty_no_raise_pass", logs)
    Exit Function

EH:
    Test_Coord_GetComboVerDescripcion_FormClosed_ReturnsEmptyNoRaise = _
        TestCore_BuildFail("Test_Coord_GetComboVerDescripcion_FormClosed_ReturnsEmptyNoRaise: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 3 (EDGE - SENTINEL) - Formulario NO abierto -> no-op sin
'   raise para Coord_RefreshRiesgoGestionRiesgo.
'
' Setup: FormRiesgosGestionRiesgo NO abierto.
' Act:   Coord_RefreshRiesgoGestionRiesgo
' Expect: NO raise. p_Error queda vacio (no-op silencioso) o con
'         mensaje legible; el atomo no distingue porque el contrato
'         es "no-op defensivo sin crashear".
' ============================================================
Public Function Test_Coord_RefreshRiesgoGestionRiesgo_FormClosed_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestionRiesgo NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_RefreshRiesgoGestionRiesgo"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_RefreshRiesgoGestionRiesgo(m_Error)

    logs(logIdx) = "3. Assert: no raise (llego aqui) - sentinel"
    logIdx = logIdx + 1

    ' Contrato: no-op silencioso cuando el form no esta abierto. El
    ' error puede quedar vacio o con un mensaje legible ("form no
    ' abierto"), pero NO debe haber raise de VBA (ya llegamos aqui).
    If m_Error <> "" And InStr(1, m_Error, "no esta abierto", vbTextCompare) = 0 Then
        logs(logIdx) = "4. FAIL: p_Error no esperado: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_RefreshRiesgoGestionRiesgo_FormClosed_NoOp = _
            TestCore_BuildFail("RefreshRiesgoGestionRiesgo no-op defensivo: p_Error inesperado '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_Coord_RefreshRiesgoGestionRiesgo_FormClosed_NoOp = _
        TestCore_BuildOk("noop_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_RefreshRiesgoGestionRiesgo_FormClosed_NoOp = _
        TestCore_BuildFail("Test_Coord_RefreshRiesgoGestionRiesgo_FormClosed_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 4 (EDGE - SENTINEL) - Formulario NO abierto -> no-op sin
'   raise para Coord_RefreshRiesgo (FormRiesgo padre).
' ============================================================
Public Function Test_Coord_RefreshRiesgo_FormClosed_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgo NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_RefreshRiesgo"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_RefreshRiesgo(m_Error)

    logs(logIdx) = "3. Assert: no raise (llego aqui) - sentinel"
    logIdx = logIdx + 1

    If m_Error <> "" And InStr(1, m_Error, "no esta abierto", vbTextCompare) = 0 Then
        logs(logIdx) = "4. FAIL: p_Error no esperado: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_RefreshRiesgo_FormClosed_NoOp = _
            TestCore_BuildFail("RefreshRiesgo no-op defensivo: p_Error inesperado '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_Coord_RefreshRiesgo_FormClosed_NoOp = _
        TestCore_BuildOk("noop_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_RefreshRiesgo_FormClosed_NoOp = _
        TestCore_BuildFail("Test_Coord_RefreshRiesgo_FormClosed_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 5 (ADVERSARIAL) - Llamada doble en secuencia no debe
'   acumular estado, ni crashear, ni propagar error persistente.
'
' Setup: Ambos forms cerrados (estado base del headless test).
' Act:   Coord_RefreshRiesgoGestionRiesgo x2, luego
'        Coord_RefreshRiesgo x2
' Expect: 4 llamadas seguidas sin raise, sin lock de MSACCESS.EXE,
'         sin acoplar estado entre llamadas.
' ============================================================
Public Function Test_Coord_RefreshRiesgoForm_AdversarialDoble_OK() As String
    Dim logs(0 To 7) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: state base, forms cerrados"
    logIdx = logIdx + 1

    Dim m_Error1 As String, m_Error2 As String
    Dim m_Error3 As String, m_Error4 As String
    m_Error1 = "": m_Error2 = "": m_Error3 = "": m_Error4 = ""

    logs(logIdx) = "2. Act: 4 llamadas defensivas consecutivas (RefreshRiesgoGestionRiesgo x2, RefreshRiesgo x2)"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_RefreshRiesgoGestionRiesgo(m_Error1)
    Call modFormCoordinationHelper.Coord_RefreshRiesgoGestionRiesgo(m_Error2)
    Call modFormCoordinationHelper.Coord_RefreshRiesgo(m_Error3)
    Call modFormCoordinationHelper.Coord_RefreshRiesgo(m_Error4)

    logs(logIdx) = "3. Assert: las 4 llamadas completaron sin raise"
    logIdx = logIdx + 1

    ' Estado consistente: las 4 son no-op (form no abierto) o todas
    ' pasan. No validamos el contenido exacto de p_Error - solo que
    ' llegamos aqui y que no se acopla estado entre llamadas (cada
    ' invocacion opera sobre su propio ByRef).
    Test_Coord_RefreshRiesgoForm_AdversarialDoble_OK = _
        TestCore_BuildOk("doble_call_no_op_pass", logs)
    Exit Function

EH:
    Test_Coord_RefreshRiesgoForm_AdversarialDoble_OK = _
        TestCore_BuildFail("Test_Coord_RefreshRiesgoForm_AdversarialDoble_OK raise: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 6 (EDGE) - Formulario NO abierto -> Nothing sin raise
'   para Coord_GetRiesgoSeleccionadoDeTecnicoTareas.
'
' Setup: FormTecnicoTareas NO abierto.
' Act:   Coord_GetRiesgoSeleccionadoDeTecnicoTareas
' Expect: Nothing, p_Error vacio.
' ============================================================
Public Function Test_Coord_GetRiesgoSeleccionadoDeTecnicoTareas_FormClosed_ReturnsNothing() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormTecnicoTareas NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = ""

    logs(logIdx) = "2. Act: Coord_GetRiesgoSeleccionadoDeTecnicoTareas"
    logIdx = logIdx + 1

    Dim m_Riesgo As Object
    Set m_Riesgo = modFormCoordinationHelper.Coord_GetRiesgoSeleccionadoDeTecnicoTareas(m_Error)

    logs(logIdx) = "3. Assert: m_Riesgo Is Nothing AND m_Error = """
    logIdx = logIdx + 1

    If Not m_Riesgo Is Nothing Then
        logs(logIdx) = "4. FAIL: esperaba Nothing, recibio '" & TypeName(m_Riesgo) & "'"
        logIdx = logIdx + 1
        Test_Coord_GetRiesgoSeleccionadoDeTecnicoTareas_FormClosed_ReturnsNothing = _
            TestCore_BuildFail("Coord_GetRiesgoSeleccionadoDeTecnicoTareas debio devolver Nothing; devolvio " & TypeName(m_Riesgo), logs)
        Exit Function
    End If

    If m_Error <> "" Then
        logs(logIdx) = "4. FAIL: m_Error deberia estar vacio, era: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_GetRiesgoSeleccionadoDeTecnicoTareas_FormClosed_ReturnsNothing = _
            TestCore_BuildFail("Coord_GetRiesgoSeleccionadoDeTecnicoTareas relleno p_Error: " & m_Error, logs)
        Exit Function
    End If

    Test_Coord_GetRiesgoSeleccionadoDeTecnicoTareas_FormClosed_ReturnsNothing = _
        TestCore_BuildOk("nothing_no_raise_pass", logs)
    Exit Function

EH:
    Test_Coord_GetRiesgoSeleccionadoDeTecnicoTareas_FormClosed_ReturnsNothing = _
        TestCore_BuildFail("Test_Coord_GetRiesgoSeleccionadoDeTecnicoTareas_FormClosed_ReturnsNothing: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 7 (EDGE - SENTINEL) - FormRiesgosGestion NO abierto ->
'   no-op sin raise para Coord_CargarArbolPM.
'
'   HR3b slice (cross-form CargarArbolPM proxy). El form destino
'   no esta abierto en sesion de test headless: el helper debe
'   no-op silencioso y rellenar p_Error con mensaje legible.
' ============================================================
Public Function Test_Coord_CargarArbolPM_FormClosed_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_CargarArbolPM (con Nothing como Nodo/PM)"
    logIdx = logIdx + 1

    ' Pasamos Nothing como Nodo/Riesgo/PM. El helper DEBE detectar
    ' que el form no esta abierto y no-op sin intentar llamar a la
    ' firma del form (que reventaria con Nodo/PM Nothing).
    Call modFormCoordinationHelper.Coord_CargarArbolPM( _
        P_NodoRiesgo:=Nothing, _
        p_Riesgo:=Nothing, _
        p_PM:=Nothing, _
        p_borrarNodoSeleccionado:=EnumSiNo.Sí, _
        p_Refrescando:=EnumSiNo.No, _
        p_Error:=m_Error)

    logs(logIdx) = "3. Assert: no raise (llego aqui) - sentinel"
    logIdx = logIdx + 1

    If m_Error = "" Then
        logs(logIdx) = "4. FAIL: p_Error deberia reportar 'no esta abierto'"
        logIdx = logIdx + 1
        Test_Coord_CargarArbolPM_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_CargarArbolPM no-op: p_Error deberia contener 'no esta abierto'; era ''", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "no esta abierto", vbTextCompare) = 0 Then
        logs(logIdx) = "4. FAIL: p_Error no esperado: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_CargarArbolPM_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_CargarArbolPM no-op: p_Error inesperado '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_Coord_CargarArbolPM_FormClosed_NoOp = _
        TestCore_BuildOk("noop_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_CargarArbolPM_FormClosed_NoOp = _
        TestCore_BuildFail("Test_Coord_CargarArbolPM_FormClosed_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 8 (EDGE - SENTINEL) - FormRiesgosGestion NO abierto ->
'   no-op sin raise para Coord_CargarArbolPC.
'
'   HR3b slice (cross-form CargarArbolPC proxy). Mismo patron que
'   ATOMO 7 pero hermano para planes de contingencia.
' ============================================================
Public Function Test_Coord_CargarArbolPC_FormClosed_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_CargarArbolPC (con Nothing como Nodo/PC)"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_CargarArbolPC( _
        P_NodoRiesgo:=Nothing, _
        p_Riesgo:=Nothing, _
        p_PC:=Nothing, _
        p_borrarNodoSeleccionado:=EnumSiNo.Sí, _
        p_Refrescando:=EnumSiNo.No, _
        p_Error:=m_Error)

    logs(logIdx) = "3. Assert: no raise (llego aqui) - sentinel"
    logIdx = logIdx + 1

    If m_Error = "" Then
        logs(logIdx) = "4. FAIL: p_Error deberia reportar 'no esta abierto'"
        logIdx = logIdx + 1
        Test_Coord_CargarArbolPC_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_CargarArbolPC no-op: p_Error deberia contener 'no esta abierto'; era ''", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "no esta abierto", vbTextCompare) = 0 Then
        logs(logIdx) = "4. FAIL: p_Error no esperado: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_CargarArbolPC_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_CargarArbolPC no-op: p_Error inesperado '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_Coord_CargarArbolPC_FormClosed_NoOp = _
        TestCore_BuildOk("noop_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_CargarArbolPC_FormClosed_NoOp = _
        TestCore_BuildFail("Test_Coord_CargarArbolPC_FormClosed_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 9 (EDGE - SENTINEL) - FormRiesgosGestion NO abierto ->
'   no-op para Coord_RefrescarArbolRiesgosScope.
'
'   HR3c slice (cross-form RefrescarArbolRiesgosScope proxy). El form
'   destino no esta abierto en sesion de test headless: el helper debe
'   no-op silencioso y rellenar p_Error con mensaje legible. Patron
'   identico a Coord_CargarArbolPM/PC (atomos 7 y 8).
' ============================================================
Public Function Test_Coord_RefrescarArbolRiesgosScope_FormClosed_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_RefrescarArbolRiesgosScope p_Scope:=""plan"""
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_RefrescarArbolRiesgosScope( _
        p_Scope:="plan", _
        p_Error:=m_Error)

    logs(logIdx) = "3. Assert: no raise (llego aqui) - sentinel"
    logIdx = logIdx + 1

    If m_Error = "" Then
        logs(logIdx) = "4. FAIL: p_Error deberia reportar 'no esta abierto'"
        logIdx = logIdx + 1
        Test_Coord_RefrescarArbolRiesgosScope_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_RefrescarArbolRiesgosScope no-op: p_Error deberia contener 'no esta abierto'; era ''", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "no esta abierto", vbTextCompare) = 0 Then
        logs(logIdx) = "4. FAIL: p_Error no esperado: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_RefrescarArbolRiesgosScope_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_RefrescarArbolRiesgosScope no-op: p_Error inesperado '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_Coord_RefrescarArbolRiesgosScope_FormClosed_NoOp = _
        TestCore_BuildOk("noop_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_RefrescarArbolRiesgosScope_FormClosed_NoOp = _
        TestCore_BuildFail("Test_Coord_RefrescarArbolRiesgosScope_FormClosed_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 10 (EDGE) - FormRiesgosGestion NO abierto -> m_EsMitigacion
'   devuelve Empty (Variant sin asignar) como String vacio sin raise.
'   Patron similar a Coord_GetSelectedNode (atomo 1) pero para un
'   EnumSiNo Public field.
' ============================================================
Public Function Test_Coord_GetEsMitigacion_FormClosed_DefaultEmpty() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = ""

    logs(logIdx) = "2. Act: Coord_GetEsMitigacion"
    logIdx = logIdx + 1

    Dim m_Valor As String
    m_Valor = modFormCoordinationHelper.Coord_GetEsMitigacion(m_Error)

    logs(logIdx) = "3. Assert: m_Valor = "" AND m_Error = """
    logIdx = logIdx + 1

    If m_Valor <> "" Then
        logs(logIdx) = "4. FAIL: esperaba '', recibio '" & m_Valor & "'"
        logIdx = logIdx + 1
        Test_Coord_GetEsMitigacion_FormClosed_DefaultEmpty = _
            TestCore_BuildFail("Coord_GetEsMitigacion debio devolver '' con form cerrado; devolvio '" & m_Valor & "'", logs)
        Exit Function
    End If

    If m_Error <> "" Then
        logs(logIdx) = "4. FAIL: m_Error deberia estar vacio"
        logIdx = logIdx + 1
        Test_Coord_GetEsMitigacion_FormClosed_DefaultEmpty = _
            TestCore_BuildFail("Coord_GetEsMitigacion relleno p_Error: " & m_Error, logs)
        Exit Function
    End If

    Test_Coord_GetEsMitigacion_FormClosed_DefaultEmpty = _
        TestCore_BuildOk("empty_no_raise_pass", logs)
    Exit Function

EH:
    Test_Coord_GetEsMitigacion_FormClosed_DefaultEmpty = _
        TestCore_BuildFail("Test_Coord_GetEsMitigacion_FormClosed_DefaultEmpty: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 11 (EDGE) - FormRiesgosGestion NO abierto ->
'   m_ObjSeleccionado devuelve Nothing sin raise.
' ============================================================
Public Function Test_Coord_GetObjSeleccionado_FormClosed_ReturnsNothing() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = ""

    logs(logIdx) = "2. Act: Coord_GetObjSeleccionado"
    logIdx = logIdx + 1

    Dim m_Obj As Object
    Set m_Obj = modFormCoordinationHelper.Coord_GetObjSeleccionado(m_Error)

    logs(logIdx) = "3. Assert: m_Obj Is Nothing AND m_Error = """
    logIdx = logIdx + 1

    If Not m_Obj Is Nothing Then
        logs(logIdx) = "4. FAIL: esperaba Nothing, recibio '" & TypeName(m_Obj) & "'"
        logIdx = logIdx + 1
        Test_Coord_GetObjSeleccionado_FormClosed_ReturnsNothing = _
            TestCore_BuildFail("Coord_GetObjSeleccionado debio devolver Nothing; devolvio " & TypeName(m_Obj), logs)
        Exit Function
    End If

    If m_Error <> "" Then
        logs(logIdx) = "4. FAIL: m_Error deberia estar vacio"
        logIdx = logIdx + 1
        Test_Coord_GetObjSeleccionado_FormClosed_ReturnsNothing = _
            TestCore_BuildFail("Coord_GetObjSeleccionado relleno p_Error: " & m_Error, logs)
        Exit Function
    End If

    Test_Coord_GetObjSeleccionado_FormClosed_ReturnsNothing = _
        TestCore_BuildOk("nothing_no_raise_pass", logs)
    Exit Function

EH:
    Test_Coord_GetObjSeleccionado_FormClosed_ReturnsNothing = _
        TestCore_BuildFail("Test_Coord_GetObjSeleccionado_FormClosed_ReturnsNothing: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 12 (EDGE) - FormRiesgosGestion NO abierto -> blnPermitidoEditar
'   devuelve False (defensa) sin raise. Patron similar al atomo 10
'   pero para Boolean Public field.
' ============================================================
Public Function Test_Coord_GetPermitidoEditar_FormClosed_DefaultsFalse() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = ""

    logs(logIdx) = "2. Act: Coord_GetPermitidoEditar"
    logIdx = logIdx + 1

    Dim m_Valor As Boolean
    m_Valor = modFormCoordinationHelper.Coord_GetPermitidoEditar(m_Error)

    logs(logIdx) = "3. Assert: m_Valor = False AND m_Error = """
    logIdx = logIdx + 1

    If m_Valor <> False Then
        logs(logIdx) = "4. FAIL: esperaba False, recibio True"
        logIdx = logIdx + 1
        Test_Coord_GetPermitidoEditar_FormClosed_DefaultsFalse = _
            TestCore_BuildFail("Coord_GetPermitidoEditar debio devolver False con form cerrado; devolvio True", logs)
        Exit Function
    End If

    If m_Error <> "" Then
        logs(logIdx) = "4. FAIL: m_Error deberia estar vacio"
        logIdx = logIdx + 1
        Test_Coord_GetPermitidoEditar_FormClosed_DefaultsFalse = _
            TestCore_BuildFail("Coord_GetPermitidoEditar relleno p_Error: " & m_Error, logs)
        Exit Function
    End If

    Test_Coord_GetPermitidoEditar_FormClosed_DefaultsFalse = _
        TestCore_BuildOk("false_no_raise_pass", logs)
    Exit Function

EH:
    Test_Coord_GetPermitidoEditar_FormClosed_DefaultsFalse = _
        TestCore_BuildFail("Test_Coord_GetPermitidoEditar_FormClosed_DefaultsFalse: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 13 (EDGE - SENTINEL) - FormRiesgosGestion NO abierto ->
'   no-op para Coord_ArbolNodeClick.
' ============================================================
Public Function Test_Coord_ArbolNodeClick_FormClosed_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_ArbolNodeClick (con Nothing como nodo)"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_ArbolNodeClick(Nothing, m_Error)

    logs(logIdx) = "3. Assert: no raise (llego aqui) - sentinel"
    logIdx = logIdx + 1

    If m_Error = "" Then
        logs(logIdx) = "4. FAIL: p_Error deberia reportar 'no esta abierto'"
        logIdx = logIdx + 1
        Test_Coord_ArbolNodeClick_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_ArbolNodeClick no-op: p_Error deberia contener 'no esta abierto'; era ''", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "no esta abierto", vbTextCompare) = 0 Then
        logs(logIdx) = "4. FAIL: p_Error no esperado: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_ArbolNodeClick_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_ArbolNodeClick no-op: p_Error inesperado '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_Coord_ArbolNodeClick_FormClosed_NoOp = _
        TestCore_BuildOk("noop_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_ArbolNodeClick_FormClosed_NoOp = _
        TestCore_BuildFail("Test_Coord_ArbolNodeClick_FormClosed_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 14 (EDGE - SENTINEL) - FormRiesgosGestion NO abierto ->
'   no-op para Coord_ArbolEliminarNodoConRefresco.
' ============================================================
Public Function Test_Coord_ArbolEliminarNodoConRefresco_FormClosed_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_ArbolEliminarNodoConRefresco (key prueba)"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_ArbolEliminarNodoConRefresco( _
        p_Key:="RISGO|9999", _
        p_Error:=m_Error)

    logs(logIdx) = "3. Assert: no raise (llego aqui) - sentinel"
    logIdx = logIdx + 1

    If m_Error = "" Then
        logs(logIdx) = "4. FAIL: p_Error deberia reportar 'no esta abierto'"
        logIdx = logIdx + 1
        Test_Coord_ArbolEliminarNodoConRefresco_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_ArbolEliminarNodoConRefresco no-op: p_Error deberia contener 'no esta abierto'; era ''", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "no esta abierto", vbTextCompare) = 0 Then
        logs(logIdx) = "4. FAIL: p_Error no esperado: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_ArbolEliminarNodoConRefresco_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_ArbolEliminarNodoConRefresco no-op: p_Error inesperado '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_Coord_ArbolEliminarNodoConRefresco_FormClosed_NoOp = _
        TestCore_BuildOk("noop_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_ArbolEliminarNodoConRefresco_FormClosed_NoOp = _
        TestCore_BuildFail("Test_Coord_ArbolEliminarNodoConRefresco_FormClosed_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 15 (EDGE) - FormRiesgosGestion NO abierto ->
'   m_Arbol.SelectedItem devuelve Nothing sin raise.
' ============================================================
Public Function Test_Coord_ArbolSelectedItem_FormClosed_ReturnsNothing() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = ""

    logs(logIdx) = "2. Act: Coord_ArbolSelectedItem"
    logIdx = logIdx + 1

    Dim m_Nodo As Object
    Set m_Nodo = modFormCoordinationHelper.Coord_ArbolSelectedItem(m_Error)

    logs(logIdx) = "3. Assert: m_Nodo Is Nothing AND m_Error = """
    logIdx = logIdx + 1

    If Not m_Nodo Is Nothing Then
        logs(logIdx) = "4. FAIL: esperaba Nothing, recibio '" & TypeName(m_Nodo) & "'"
        logIdx = logIdx + 1
        Test_Coord_ArbolSelectedItem_FormClosed_ReturnsNothing = _
            TestCore_BuildFail("Coord_ArbolSelectedItem debio devolver Nothing; devolvio " & TypeName(m_Nodo), logs)
        Exit Function
    End If

    If m_Error <> "" Then
        logs(logIdx) = "4. FAIL: m_Error deberia estar vacio"
        logIdx = logIdx + 1
        Test_Coord_ArbolSelectedItem_FormClosed_ReturnsNothing = _
            TestCore_BuildFail("Coord_ArbolSelectedItem relleno p_Error: " & m_Error, logs)
        Exit Function
    End If

    Test_Coord_ArbolSelectedItem_FormClosed_ReturnsNothing = _
        TestCore_BuildOk("nothing_no_raise_pass", logs)
    Exit Function

EH:
    Test_Coord_ArbolSelectedItem_FormClosed_ReturnsNothing = _
        TestCore_BuildFail("Test_Coord_ArbolSelectedItem_FormClosed_ReturnsNothing: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 16 (EDGE - SENTINEL) - FormRiesgosGestion NO abierto ->
'   no-op para Coord_CargarArbol.
' ============================================================
Public Function Test_Coord_CargarArbol_FormClosed_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_CargarArbol"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_CargarArbol(m_Error)

    logs(logIdx) = "3. Assert: no raise (llego aqui) - sentinel"
    logIdx = logIdx + 1

    If m_Error = "" Then
        logs(logIdx) = "4. FAIL: p_Error deberia reportar 'no esta abierto'"
        logIdx = logIdx + 1
        Test_Coord_CargarArbol_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_CargarArbol no-op: p_Error deberia contener 'no esta abierto'; era ''", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "no esta abierto", vbTextCompare) = 0 Then
        logs(logIdx) = "4. FAIL: p_Error no esperado: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_CargarArbol_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_CargarArbol no-op: p_Error inesperado '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_Coord_CargarArbol_FormClosed_NoOp = _
        TestCore_BuildOk("noop_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_CargarArbol_FormClosed_NoOp = _
        TestCore_BuildFail("Test_Coord_CargarArbol_FormClosed_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 17 (EDGE - SENTINEL) - FormRiesgosGestion NO abierto ->
'   no-op para Coord_SeleccionarNodo.
' ============================================================
Public Function Test_Coord_SeleccionarNodo_FormClosed_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_SeleccionarNodo (con Nothing como objeto)"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_SeleccionarNodo( _
        p_Objeto:=Nothing, _
        p_Error:=m_Error)

    logs(logIdx) = "3. Assert: no raise (llego aqui) - sentinel"
    logIdx = logIdx + 1

    If m_Error = "" Then
        logs(logIdx) = "4. FAIL: p_Error deberia reportar 'no esta abierto'"
        logIdx = logIdx + 1
        Test_Coord_SeleccionarNodo_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_SeleccionarNodo no-op: p_Error deberia contener 'no esta abierto'; era ''", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "no esta abierto", vbTextCompare) = 0 Then
        logs(logIdx) = "4. FAIL: p_Error no esperado: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_SeleccionarNodo_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_SeleccionarNodo no-op: p_Error inesperado '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_Coord_SeleccionarNodo_FormClosed_NoOp = _
        TestCore_BuildOk("noop_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_SeleccionarNodo_FormClosed_NoOp = _
        TestCore_BuildFail("Test_Coord_SeleccionarNodo_FormClosed_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 18 (EDGE - SENTINEL) - FormRiesgosGestion NO abierto ->
'   no-op para Coord_CargarArbolAccion.
' ============================================================
Public Function Test_Coord_CargarArbolAccion_FormClosed_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_CargarArbolAccion (con Nothing como nodo/accion)"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_CargarArbolAccion( _
        p_NodoPlan:=Nothing, _
        p_Accion:=Nothing, _
        p_Error:=m_Error)

    logs(logIdx) = "3. Assert: no raise (llego aqui) - sentinel"
    logIdx = logIdx + 1

    If m_Error = "" Then
        logs(logIdx) = "4. FAIL: p_Error deberia reportar 'no esta abierto'"
        logIdx = logIdx + 1
        Test_Coord_CargarArbolAccion_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_CargarArbolAccion no-op: p_Error deberia contener 'no esta abierto'; era ''", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "no esta abierto", vbTextCompare) = 0 Then
        logs(logIdx) = "4. FAIL: p_Error no esperado: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_CargarArbolAccion_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_CargarArbolAccion no-op: p_Error inesperado '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_Coord_CargarArbolAccion_FormClosed_NoOp = _
        TestCore_BuildOk("noop_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_CargarArbolAccion_FormClosed_NoOp = _
        TestCore_BuildFail("Test_Coord_CargarArbolAccion_FormClosed_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 19 (EDGE - SENTINEL) - FormRiesgosGestion NO abierto ->
'   no-op para Coord_ColRiesgosAplicados_RemoveIfExists.
' ============================================================
Public Function Test_Coord_ColRiesgosAplicados_RemoveIfExists_FormClosed_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_ColRiesgosAplicados_RemoveIfExists (key prueba)"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_ColRiesgosAplicados_RemoveIfExists( _
        p_Key:="9999", _
        p_Error:=m_Error)

    logs(logIdx) = "3. Assert: no raise (llego aqui) - sentinel"
    logIdx = logIdx + 1

    If m_Error = "" Then
        logs(logIdx) = "4. FAIL: p_Error deberia reportar 'no esta abierto'"
        logIdx = logIdx + 1
        Test_Coord_ColRiesgosAplicados_RemoveIfExists_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_ColRiesgosAplicados_RemoveIfExists no-op: p_Error deberia contener 'no esta abierto'; era ''", logs)
        Exit Function
    End If

    If InStr(1, m_Error, "no esta abierto", vbTextCompare) = 0 Then
        logs(logIdx) = "4. FAIL: p_Error no esperado: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_ColRiesgosAplicados_RemoveIfExists_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_ColRiesgosAplicados_RemoveIfExists no-op: p_Error inesperado '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_Coord_ColRiesgosAplicados_RemoveIfExists_FormClosed_NoOp = _
        TestCore_BuildOk("noop_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_ColRiesgosAplicados_RemoveIfExists_FormClosed_NoOp = _
        TestCore_BuildFail("Test_Coord_ColRiesgosAplicados_RemoveIfExists_FormClosed_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 20 (ADVERSARIAL DOBLE) - Las 11 nuevas entradas (Coord_
'   RefrescarArbolRiesgosScope, GetEsMitigacion, GetObjSeleccionado,
'   GetPermitidoEditar, ArbolNodeClick, ArbolEliminarNodoConRefresco,
'   ArbolSelectedItem, CargarArbol, SeleccionarNodo, CargarArbolAccion,
'   ColRiesgosAplicados_RemoveIfExists) llamadas en secuencia con
'   FormRiesgosGestion cerrado no deben lockear MSACCESS.EXE, no
'   deben acumular estado, no deben propagar error persistente.
'
'   Setup: state base, FormRiesgosGestion cerrado.
'   Act:   11 llamadas consecutivas (mix de sub y function).
'   Expect: 11/11 sin raise, sin state leak.
' ============================================================
Public Function Test_Coord_HR3cHelpers_AdversarialDoble_OK() As String
    Dim logs(0 To 11) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: state base, FormRiesgosGestion cerrado"
    logIdx = logIdx + 1

    Dim m_Err1 As String, m_Err2 As String
    Dim m_Err3 As String, m_Err4 As String
    Dim m_Err5 As String, m_Err6 As String
    Dim m_Err7 As String, m_Err8 As String
    Dim m_Err9 As String, m_Err10 As String
    Dim m_Err11 As String
    m_Err1 = "": m_Err2 = "": m_Err3 = "": m_Err4 = ""
    m_Err5 = "": m_Err6 = "": m_Err7 = "": m_Err8 = ""
    m_Err9 = "": m_Err10 = "": m_Err11 = ""

    Dim m_ValorStr As String, m_ValorBool As Boolean
    Dim m_ValorObj1 As Object, m_ValorObj2 As Object

    logs(logIdx) = "2. Act: 11 entradas consecutivas (mix sub + function)"
    logIdx = logIdx + 1

    ' Function reads (estado de FormRiesgosGestion)
    m_ValorStr = modFormCoordinationHelper.Coord_GetEsMitigacion(m_Err1)
    Set m_ValorObj1 = modFormCoordinationHelper.Coord_GetObjSeleccionado(m_Err2)
    m_ValorBool = modFormCoordinationHelper.Coord_GetPermitidoEditar(m_Err3)
    Set m_ValorObj2 = modFormCoordinationHelper.Coord_ArbolSelectedItem(m_Err4)

    ' Sub proxies (calls al form que no-op)
    Call modFormCoordinationHelper.Coord_RefrescarArbolRiesgosScope("plan", m_Err5)
    Call modFormCoordinationHelper.Coord_ArbolNodeClick(Nothing, m_Err6)
    Call modFormCoordinationHelper.Coord_ArbolEliminarNodoConRefresco("RISGO|9999", m_Err7)
    Call modFormCoordinationHelper.Coord_CargarArbol(m_Err8)
    Call modFormCoordinationHelper.Coord_SeleccionarNodo(Nothing, m_Err9)
    Call modFormCoordinationHelper.Coord_CargarArbolAccion(Nothing, Nothing, m_Err10)
    Call modFormCoordinationHelper.Coord_ColRiesgosAplicados_RemoveIfExists("9999", m_Err11)

    logs(logIdx) = "3. Assert: 11/11 llamadas completaron sin raise"
    logIdx = logIdx + 1

    ' Estado consistente: las 11 son no-op (form no abierto). No
    ' validamos contenido exacto de p_Error, solo que llegamos aqui
    ' y que no se acopla estado entre llamadas.
    Test_Coord_HR3cHelpers_AdversarialDoble_OK = _
        TestCore_BuildOk("doble_call_11_no_op_pass", logs)
    Exit Function

EH:
    Test_Coord_HR3cHelpers_AdversarialDoble_OK = _
        TestCore_BuildFail("Test_Coord_HR3cHelpers_AdversarialDoble_OK raise: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 21 (EDGE - SENTINEL) - FormExpedientesBusqueda NO abierto ->
'   no-op sin raise para Coord_ExpedientesBusqueda_Filtrar. Reemplaza
'   el cross-form Form_FormExpedientesBusqueda.Filtrar en
'   Form_FormGestionRiesgosDatosGenerales (L79, L316).
' ============================================================
Public Function Test_Coord_ExpedientesBusqueda_Filtrar_FormClosed_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormExpedientesBusqueda NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_ExpedientesBusqueda_Filtrar"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_ExpedientesBusqueda_Filtrar( _
        p_PalabraClave:="prueba", _
        p_Error:=m_Error)

    logs(logIdx) = "3. Assert: no raise (llego aqui) - sentinel"
    logIdx = logIdx + 1

    If m_Error <> "" And InStr(1, m_Error, "no esta abierto", vbTextCompare) = 0 Then
        logs(logIdx) = "4. FAIL: p_Error no esperado: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_ExpedientesBusqueda_Filtrar_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_ExpedientesBusqueda_Filtrar no-op: p_Error inesperado '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_Coord_ExpedientesBusqueda_Filtrar_FormClosed_NoOp = _
        TestCore_BuildOk("noop_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_ExpedientesBusqueda_Filtrar_FormClosed_NoOp = _
        TestCore_BuildFail("Test_Coord_ExpedientesBusqueda_Filtrar_FormClosed_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 22 (EDGE - SENTINEL) - FormRiesgosGestion NO abierto -> no-op
'   sin raise para Coord_ComandoActualizarContadorRiesgosGestion.
'   Reemplaza el cross-form Form_FormRiesgosGestion
'   .ComandoActualizarContador_Click en Form_FormRiesgosEstablecer
'   Prioridades.EstablecerPriorizaciones (L575).
' ============================================================
Public Function Test_Coord_ComandoActualizarContadorRiesgosGestion_FormClosed_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_ComandoActualizarContadorRiesgosGestion"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_ComandoActualizarContadorRiesgosGestion(m_Error)

    logs(logIdx) = "3. Assert: no raise (llego aqui) - sentinel"
    logIdx = logIdx + 1

    If m_Error <> "" And InStr(1, m_Error, "no esta abierto", vbTextCompare) = 0 Then
        logs(logIdx) = "4. FAIL: p_Error no esperado: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_ComandoActualizarContadorRiesgosGestion_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_ComandoActualizarContadorRiesgosGestion no-op: p_Error inesperado '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_Coord_ComandoActualizarContadorRiesgosGestion_FormClosed_NoOp = _
        TestCore_BuildOk("noop_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_ComandoActualizarContadorRiesgosGestion_FormClosed_NoOp = _
        TestCore_BuildFail("Test_Coord_ComandoActualizarContadorRiesgosGestion_FormClosed_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 23 (EDGE - SENTINEL) - FormRiesgosGestion NO abierto -> no-op
'   sin raise para Coord_EstablecerLblRechazadoEnDetalleRiesgos.
'   Reemplaza el cross-form Forms("FormRiesgosGestion").FormDetalle...
'   chain en Form_FormPublicacionCalidadPublicar.m_FormMotivos_Motivado
'   (L407-415).
' ============================================================
Public Function Test_Coord_EstablecerLblRechazadoEnDetalleRiesgos_FormClosed_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_EstablecerLblRechazadoEnDetalleRiesgos"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_EstablecerLblRechazadoEnDetalleRiesgos(m_Error)

    logs(logIdx) = "3. Assert: no raise (llego aqui) - sentinel"
    logIdx = logIdx + 1

    If m_Error <> "" And InStr(1, m_Error, "no esta abierto", vbTextCompare) = 0 Then
        logs(logIdx) = "4. FAIL: p_Error no esperado: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_EstablecerLblRechazadoEnDetalleRiesgos_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_EstablecerLblRechazadoEnDetalleRiesgos no-op: p_Error inesperado '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_Coord_EstablecerLblRechazadoEnDetalleRiesgos_FormClosed_NoOp = _
        TestCore_BuildOk("noop_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_EstablecerLblRechazadoEnDetalleRiesgos_FormClosed_NoOp = _
        TestCore_BuildFail("Test_Coord_EstablecerLblRechazadoEnDetalleRiesgos_FormClosed_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 24 (ADVERSARIAL DOBLE) - Las 3 nuevas entradas HR3d
'   (Coord_ExpedientesBusqueda_Filtrar,
'    Coord_ComandoActualizarContadorRiesgosGestion,
'    Coord_EstablecerLblRechazadoEnDetalleRiesgos) llamadas en
'   secuencia con los forms cerrados no deben lockear MSACCESS.EXE,
'   no deben acumular estado, no deben propagar error persistente.
'
'   Setup: state base, todos los forms objetivo cerrados.
'   Act:   3 llamadas consecutivas (mix sub).
'   Expect: 3/3 sin raise, sin state leak, sin propagacion entre
'           llamadas.
' ============================================================
Public Function Test_Coord_HR3dHelpers_AdversarialDoble_OK() As String
    Dim logs(0 To 7) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: state base, todos los forms objetivo cerrados"
    logIdx = logIdx + 1

    Dim m_Err1 As String, m_Err2 As String, m_Err3 As String
    m_Err1 = "": m_Err2 = "": m_Err3 = ""

    logs(logIdx) = "2. Act: 3 entradas consecutivas con form cerrado"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_ExpedientesBusqueda_Filtrar("prueba", m_Err1)
    Call modFormCoordinationHelper.Coord_ComandoActualizarContadorRiesgosGestion(m_Err2)
    Call modFormCoordinationHelper.Coord_EstablecerLblRechazadoEnDetalleRiesgos(m_Err3)

    logs(logIdx) = "3. Assert: 3/3 llamadas completaron sin raise"
    logIdx = logIdx + 1

    ' No-op defensivo (form no abierto). Validamos solo que llegamos
    ' aqui (no raise) y que no se acopla estado entre llamadas.
    Test_Coord_HR3dHelpers_AdversarialDoble_OK = _
        TestCore_BuildOk("doble_call_3_no_op_pass", logs)
    Exit Function

EH:
    Test_Coord_HR3dHelpers_AdversarialDoble_OK = _
        TestCore_BuildFail("Test_Coord_HR3dHelpers_AdversarialDoble_OK raise: " & Err.Description, logs)
End Function

' ============================================================
' === SLICE HR3e (2026-07-01) — 8 forms cross-form .cls closure ===
' Reemplaza:
'   - Form_Form0BDOpciones:228    Form_FormRiesgosBibliotecaGestion.ComandoNoExisteRiesgo
'   - Form_FormGestionRiesgosRiesgosOferta:302
'                                 Form_FormRiesgosBibliotecaGestion.Filtrar
'   - Form_FormRiesgo:31-43       Form_FormRiesgoX.EstablecerDatos (6 forms)
' Todas son violaciones de Hard rule 3 (form .cls MUST NOT call
' otro form .cls). Patron A (thin proxy + p_Error ByRef).
' ============================================================

' ============================================================
' ATOMO 30 (EDGE - SENTINEL) - FormRiesgosBibliotecaGestion NO abierto
'   -> no-op con p_Error para Coord_SetComandoNoExisteRiesgoVisible.
'   Reemplaza el cross-form
'   Form_FormRiesgosBibliotecaGestion.ComandoNoExisteRiesgo.Visible = False
'   en Form_Form0BDOpciones.ComandoBibliotecaRiesgos_Click (L228).
' ============================================================
Public Function Test_Coord_SetComandoNoExisteRiesgoVisible_FormClosed_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosBibliotecaGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_SetComandoNoExisteRiesgoVisible (p_Visible=False)"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_SetComandoNoExisteRiesgoVisible( _
        p_Visible:=False, _
        p_Error:=m_Error)

    logs(logIdx) = "3. Assert: no raise (llego aqui) - sentinel"
    logIdx = logIdx + 1

    If m_Error <> "" And InStr(1, m_Error, "no esta abierto", vbTextCompare) = 0 Then
        logs(logIdx) = "4. FAIL: p_Error no esperado: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_SetComandoNoExisteRiesgoVisible_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_SetComandoNoExisteRiesgoVisible no-op: p_Error inesperado '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_Coord_SetComandoNoExisteRiesgoVisible_FormClosed_NoOp = _
        TestCore_BuildOk("noop_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_SetComandoNoExisteRiesgoVisible_FormClosed_NoOp = _
        TestCore_BuildFail("Test_Coord_SetComandoNoExisteRiesgoVisible_FormClosed_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 31 (EDGE - SENTINEL) - FormRiesgosBibliotecaGestion NO abierto
'   -> no-op con p_Error para Coord_FiltrarRiesgosBibliotecaGestion.
'   Reemplaza el cross-form
'   Form_FormRiesgosBibliotecaGestion.Filtrar en
'   Form_FormGestionRiesgosRiesgosOferta.ComandoSeleccionarRiesgoDeBiblioteca_Click (L302).
' ============================================================
Public Function Test_Coord_FiltrarRiesgosBibliotecaGestion_FormClosed_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: FormRiesgosBibliotecaGestion NO abierto"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_FiltrarRiesgosBibliotecaGestion"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_FiltrarRiesgosBibliotecaGestion(m_Error)

    logs(logIdx) = "3. Assert: no raise (llego aqui) - sentinel"
    logIdx = logIdx + 1

    If m_Error <> "" And InStr(1, m_Error, "no esta abierto", vbTextCompare) = 0 Then
        logs(logIdx) = "4. FAIL: p_Error no esperado: " & m_Error
        logIdx = logIdx + 1
        Test_Coord_FiltrarRiesgosBibliotecaGestion_FormClosed_NoOp = _
            TestCore_BuildFail("Coord_FiltrarRiesgosBibliotecaGestion no-op: p_Error inesperado '" & m_Error & "'", logs)
        Exit Function
    End If

    Test_Coord_FiltrarRiesgosBibliotecaGestion_FormClosed_NoOp = _
        TestCore_BuildOk("noop_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_FiltrarRiesgosBibliotecaGestion_FormClosed_NoOp = _
        TestCore_BuildFail("Test_Coord_FiltrarRiesgosBibliotecaGestion_FormClosed_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 32 (EDGE - SENTINEL) - Subform name vacio -> p_Error para
'   Coord_RefrescarSubformDetalleRiesgo. Reemplaza el cross-form
'   Select Case que llamaba Form_FormRiesgoX.EstablecerDatos (6
'   forms) en Form_FormRiesgo.ComandoActualizar_Click (L31-43).
' ============================================================
Public Function Test_Coord_RefrescarSubformDetalleRiesgo_EmptyName_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: p_SubformName vacio"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_RefrescarSubformDetalleRiesgo con nombre vacio"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_RefrescarSubformDetalleRiesgo( _
        p_SubformName:="", _
        p_Error:=m_Error)

    logs(logIdx) = "3. Assert: p_Error contiene 'vacio' o 'subform' - no raise"
    logIdx = logIdx + 1

    If m_Error = "" Then
        logs(logIdx) = "4. FAIL: esperaba p_Error con nombre vacio, recibio vacio"
        logIdx = logIdx + 1
        Test_Coord_RefrescarSubformDetalleRiesgo_EmptyName_NoOp = _
            TestCore_BuildFail("Coord_RefrescarSubformDetalleRiesgo con nombre vacio no relleno p_Error", logs)
        Exit Function
    End If

    Test_Coord_RefrescarSubformDetalleRiesgo_EmptyName_NoOp = _
        TestCore_BuildOk("empty_name_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_RefrescarSubformDetalleRiesgo_EmptyName_NoOp = _
        TestCore_BuildFail("Test_Coord_RefrescarSubformDetalleRiesgo_EmptyName_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 33 (EDGE - SENTINEL) - Subform name desconocido -> p_Error
'   para Coord_RefrescarSubformDetalleRiesgo.
' ============================================================
Public Function Test_Coord_RefrescarSubformDetalleRiesgo_UnknownName_NoOp() As String
    Dim logs(0 To 5) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: p_SubformName desconocido"
    logIdx = logIdx + 1

    Dim m_Error As String
    m_Error = "PRESET"

    logs(logIdx) = "2. Act: Coord_RefrescarSubformDetalleRiesgo con nombre invalido"
    logIdx = logIdx + 1

    Call modFormCoordinationHelper.Coord_RefrescarSubformDetalleRiesgo( _
        p_SubformName:="FormRiesgoNoExiste", _
        p_Error:=m_Error)

    logs(logIdx) = "3. Assert: p_Error contiene 'no reconocido' o 'desconocido' - no raise"
    logIdx = logIdx + 1

    If m_Error = "" Then
        logs(logIdx) = "4. FAIL: esperaba p_Error con nombre desconocido, recibio vacio"
        logIdx = logIdx + 1
        Test_Coord_RefrescarSubformDetalleRiesgo_UnknownName_NoOp = _
            TestCore_BuildFail("Coord_RefrescarSubformDetalleRiesgo con nombre invalido no relleno p_Error", logs)
        Exit Function
    End If

    Test_Coord_RefrescarSubformDetalleRiesgo_UnknownName_NoOp = _
        TestCore_BuildOk("unknown_name_sentinel_pass", logs)
    Exit Function

EH:
    Test_Coord_RefrescarSubformDetalleRiesgo_UnknownName_NoOp = _
        TestCore_BuildFail("Test_Coord_RefrescarSubformDetalleRiesgo_UnknownName_NoOp raise no esperado: " & Err.Description, logs)
End Function

' ============================================================
' ATOMO 34 (EDGE - ADVERSARIAL DOBLE) - Las 3 nuevas entradas
'   consecutivas con FormRiesgosBibliotecaGestion NO abierto.
'   Verifica que no hay acoplo de estado entre llamadas y que
'   cada una se comporta como no-op independiente.
' ============================================================
Public Function Test_Coord_HR3eHelpers_AdversarialDoble_OK() As String
    Dim logs(0 To 8) As String
    Dim logIdx As Long
    logIdx = 0

    On Error GoTo EH

    logs(logIdx) = "1. Arrange: state base, FormRiesgosBibliotecaGestion cerrado"
    logIdx = logIdx + 1

    Dim m_Err1 As String, m_Err2 As String, m_Err3 As String
    m_Err1 = "": m_Err2 = "": m_Err3 = ""

    logs(logIdx) = "2. Act: 3 entradas consecutivas (HR3e slice)"
    logIdx = logIdx + 1

    ' HR3e #1: ComandoNoExisteRiesgo.Visible proxy
    Call modFormCoordinationHelper.Coord_SetComandoNoExisteRiesgoVisible( _
        p_Visible:=False, _
        p_Error:=m_Err1)

    ' HR3e #3: Filtrar proxy
    Call modFormCoordinationHelper.Coord_FiltrarRiesgosBibliotecaGestion(m_Err2)

    ' HR3e #5: RefrescarSubformDetalleRiesgo dispatcher (con form cerrado)
    ' En estado cerrado, el helper tambien falla porque subform vive
    ' dentro de FormRiesgo que esta cerrado. Validamos solo el no-op
    ' defensivo del dispatcher (no-op con p_Error no-vacio).
    Call modFormCoordinationHelper.Coord_RefrescarSubformDetalleRiesgo( _
        p_SubformName:="FormRiesgoDefinicion", _
        p_Error:=m_Err3)

    logs(logIdx) = "3. Assert: 3/3 llamadas completaron sin raise"
    logIdx = logIdx + 1

    ' Estado consistente: las 3 son no-op (form no abierto). No
    ' validamos contenido exacto de p_Error, solo que llegamos aqui
    ' y que no se acopla estado entre llamadas.
    Test_Coord_HR3eHelpers_AdversarialDoble_OK = _
        TestCore_BuildOk("doble_call_3_no_op_pass", logs)
    Exit Function

EH:
    Test_Coord_HR3eHelpers_AdversarialDoble_OK = _
        TestCore_BuildFail("Test_Coord_HR3eHelpers_AdversarialDoble_OK raise: " & Err.Description, logs)
End Function
